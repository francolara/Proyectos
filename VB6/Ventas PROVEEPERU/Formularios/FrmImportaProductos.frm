VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmImportaProductos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Productos"
   ClientHeight    =   5370
   ClientLeft      =   2445
   ClientTop       =   2295
   ClientWidth     =   10185
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
   ScaleHeight     =   5370
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImportar 
      Caption         =   "&Importar"
      Height          =   420
      Left            =   3540
      TabIndex        =   38
      Top             =   4800
      Width           =   1365
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   5040
      TabIndex        =   37
      Top             =   4800
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10095
      Begin VB.CommandButton CmbAyudaExcel 
         Height          =   315
         Left            =   9435
         Picture         =   "FrmImportaProductos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4230
         Width           =   390
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   " Cuentas Contables "
         ForeColor       =   &H80000008&
         Height          =   1830
         Left            =   120
         TabIndex        =   21
         Top             =   2310
         Width           =   9825
         Begin VB.CommandButton CmbAyudaCtaContableRelacionada 
            Height          =   315
            Left            =   9300
            Picture         =   "FrmImportaProductos.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1275
            Width           =   390
         End
         Begin VB.CommandButton CmbAyudaCtaContableCompra 
            Height          =   315
            Left            =   9300
            Picture         =   "FrmImportaProductos.frx":0714
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   870
            Width           =   390
         End
         Begin VB.CommandButton CmbAyudaCtaContableVenta 
            Height          =   315
            Left            =   9300
            Picture         =   "FrmImportaProductos.frx":0A9E
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   450
            Width           =   390
         End
         Begin CATControls.CATTextBox TxtIdCtaContableVenta 
            Height          =   315
            Left            =   1260
            TabIndex        =   25
            Tag             =   "TCtaContable"
            Top             =   450
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
            MaxLength       =   12
            Container       =   "FrmImportaProductos.frx":0E28
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox TxtIdCtaContableCompra 
            Height          =   315
            Left            =   1260
            TabIndex        =   26
            Tag             =   "TCtaContable2"
            Top             =   870
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
            MaxLength       =   12
            Container       =   "FrmImportaProductos.frx":0E44
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox TxtIdCtaContableRelacionada 
            Height          =   315
            Left            =   1260
            TabIndex        =   27
            Tag             =   "TCtaContable_Relacionada"
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
            MaxLength       =   12
            Container       =   "FrmImportaProductos.frx":0E60
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsCtaContableVenta 
            Height          =   315
            Left            =   2190
            TabIndex        =   28
            Top             =   450
            Width           =   7065
            _ExtentX        =   12462
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
            Container       =   "FrmImportaProductos.frx":0E7C
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsCtaContableCompra 
            Height          =   315
            Left            =   2190
            TabIndex        =   29
            Top             =   870
            Width           =   7065
            _ExtentX        =   12462
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
            Container       =   "FrmImportaProductos.frx":0E98
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsCtaContableRelacionada 
            Height          =   315
            Left            =   2190
            TabIndex        =   30
            Top             =   1275
            Width           =   7065
            _ExtentX        =   12462
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
            Container       =   "FrmImportaProductos.frx":0EB4
            Vacio           =   -1  'True
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Relacionada"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   210
            TabIndex        =   33
            Top             =   1335
            Width           =   885
         End
         Begin VB.Label Venta 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Venta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   210
            TabIndex        =   32
            Top             =   495
            Width           =   435
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Compra"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   210
            TabIndex        =   31
            Top             =   915
            Width           =   555
         End
      End
      Begin VB.CommandButton CmbAyudaTipoProducto 
         Height          =   315
         Left            =   9390
         Picture         =   "FrmImportaProductos.frx":0ED0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   390
      End
      Begin VB.CommandButton CmbAyudaMarca 
         Height          =   315
         Left            =   9390
         Picture         =   "FrmImportaProductos.frx":125A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1125
         Width           =   390
      End
      Begin VB.CommandButton CmbAyudaMoneda 
         Height          =   315
         Left            =   9390
         Picture         =   "FrmImportaProductos.frx":15E4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1530
         Width           =   390
      End
      Begin VB.CommandButton CmbAyudaUM 
         Height          =   315
         Left            =   9390
         Picture         =   "FrmImportaProductos.frx":196E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1935
         Width           =   390
      End
      Begin VB.CommandButton CmbAyudaNivel 
         Height          =   315
         Left            =   9390
         Picture         =   "FrmImportaProductos.frx":1CF8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox TxtIdNivel 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         Tag             =   "TidNivelPred"
         Top             =   300
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
         Container       =   "FrmImportaProductos.frx":2082
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsNivel 
         Height          =   315
         Left            =   2295
         TabIndex        =   2
         Top             =   300
         Width           =   7065
         _ExtentX        =   12462
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
         Container       =   "FrmImportaProductos.frx":209E
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtIdTipoProducto 
         Height          =   315
         Left            =   1365
         TabIndex        =   9
         Tag             =   "TidTipoProducto"
         Top             =   720
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
         Container       =   "FrmImportaProductos.frx":20BA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsTipoProducto 
         Height          =   315
         Left            =   2295
         TabIndex        =   10
         Top             =   720
         Width           =   7065
         _ExtentX        =   12462
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
         Container       =   "FrmImportaProductos.frx":20D6
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtIdMarca 
         Height          =   315
         Left            =   1365
         TabIndex        =   11
         Tag             =   "TidMarca"
         Top             =   1125
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
         Container       =   "FrmImportaProductos.frx":20F2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsMarca 
         Height          =   315
         Left            =   2295
         TabIndex        =   12
         Top             =   1125
         Width           =   7065
         _ExtentX        =   12462
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
         Container       =   "FrmImportaProductos.frx":210E
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtIdMoneda 
         Height          =   315
         Left            =   1365
         TabIndex        =   13
         Tag             =   "TidMoneda"
         Top             =   1530
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
         Container       =   "FrmImportaProductos.frx":212A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsMoneda 
         Height          =   315
         Left            =   2295
         TabIndex        =   14
         Top             =   1530
         Width           =   7065
         _ExtentX        =   12462
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
         Container       =   "FrmImportaProductos.frx":2146
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtIdUM 
         Height          =   315
         Left            =   1365
         TabIndex        =   15
         Tag             =   "TidUMCompra"
         Top             =   1935
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
         Container       =   "FrmImportaProductos.frx":2162
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsUM 
         Height          =   315
         Left            =   2295
         TabIndex        =   16
         Top             =   1935
         Width           =   7065
         _ExtentX        =   12462
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
         Container       =   "FrmImportaProductos.frx":217E
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsExcel 
         Height          =   330
         Left            =   1380
         TabIndex        =   35
         Top             =   4230
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial Narrow"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         Container       =   "FrmImportaProductos.frx":219A
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo Excel"
         Height          =   255
         Left            =   180
         TabIndex        =   36
         Top             =   4305
         Width           =   1155
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "U.M."
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   20
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   19
         Top             =   1575
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   18
         Top             =   1170
         Width           =   450
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   17
         Top             =   765
         Width           =   990
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nivel"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   345
      End
   End
   Begin MSComDlg.CommonDialog CdExcel 
      Left            =   390
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmImportaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbAyudaCtaContableCompra_Click()
On Error GoTo Err
Dim StrMsgError                     As String
Dim CId                             As String

    mostrarAyudaTextoPlanCuentas strcnConta, "PLANCUENTAS", CId, "", "", "2011"
    
    TxtIdCtaContableCompra.Text = CId
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaCtaContableRelacionada_Click()
On Error GoTo Err
Dim StrMsgError                     As String
Dim CId                             As String

    mostrarAyudaTextoPlanCuentas strcnConta, "PLANCUENTAS", CId, "", "", "2011"
    
    TxtIdCtaContableRelacionada.Text = CId
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaCtaContableVenta_Click()
On Error GoTo Err
Dim StrMsgError                     As String
Dim CId                             As String

    mostrarAyudaTextoPlanCuentas strcnConta, "PLANCUENTAS", CId, "", "", "2011"
    
    TxtIdCtaContableVenta.Text = CId
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaExcel_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    CdExcel.Filter = "Microsoft Excel (*.xls)|*.xls"
    CdExcel.ShowOpen
    TxtGlsExcel.Text = CdExcel.FileName
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaMarca_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    mostrarAyuda "MARCA", TxtIdMarca, TxtGlsMarca
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaMoneda_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    mostrarAyuda "MONEDA", TxtIdMoneda, TxtGlsMoneda
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaNivel_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    mostrarAyuda "NIVEL", TxtIdNivel, TxtGlsNivel, "And IdTipoNivel In(Select IdTipoNivel From TiposNiveles Where IdEmpresa = '" & glsEmpresa & "' And Peso = " & glsNumNiveles & ")"
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaTipoProducto_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    mostrarAyuda "TIPOPRODUCTO", TxtIdTipoProducto, TxtGlsTipoProducto
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaUM_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    mostrarAyuda "UM", TxtIdUM, TxtGlsUM
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdImportar_Click()
On Error GoTo Err
Dim StrMsgError                     As String
Dim xl                              As New Excel.Application
Dim wb                              As Workbook
Dim shRxn                           As Worksheet
Dim NFil                            As Integer
Dim NItem                           As Integer
Dim indTrans                        As Boolean
Dim CSqlC                           As String
Dim CIdProducto                     As String
Dim CNumero                         As String
Dim RsC                             As New ADODB.Recordset
Dim NVVUnit                         As Double
Dim NIGVUnit                        As Double
Dim NPVUnit                         As Double
Dim CIdLista                        As String

    If TxtGlsNivel.Text = "" Then StrMsgError = "Ingrese el Nivel.": GoTo Err
    If TxtGlsTipoProducto.Text = "" Then StrMsgError = "Ingrese el Tipo de Producto.": GoTo Err
    If TxtGlsMarca.Text = "" Then StrMsgError = "Ingrese la Marca.": GoTo Err
    If TxtGlsMoneda.Text = "" Then StrMsgError = "Ingrese la Moneda.": GoTo Err
    If TxtGlsUM.Text = "" Then StrMsgError = "Ingrese la Unidad de Medida.": GoTo Err
    If TxtGlsCtaContableVenta.Text = "" Then StrMsgError = "Ingrese la Cuenta Contable de Ventas.": GoTo Err
    CIdLista = leeParametro("LISTAVENTAS")
    MousePointer = 13
    indTrans = False
    
    Set xl = New Excel.Application
    Set wb = xl.Workbooks.Open(TxtGlsExcel.Text)
    
    xl.Cells.Select
    With xl.Selection
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    wb.Worksheets(1).Select
    NFil = 2
    
    xl.Visible = False
    
    Cn.BeginTrans
    indTrans = True
    
    Do While Len(xl.Cells(NFil, 1).Value) > 0
        
        CIdProducto = ""
        
        CSqlC = "Select A.IdProducto " & _
                "From Productos A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdFabricante = '" & xl.Cells(NFil, 1).Value & "'"
        AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not RsC.EOF Then
            CIdProducto = "" & RsC.Fields("IdProducto")
        End If
        RsC.Close: Set RsC = Nothing
        
        If CIdProducto = "" Then
            
            CIdProducto = GeneraCorrelativoAnoMes("Productos", "IdProducto", True)
            
            CSqlC = "Insert Into Productos(IdProducto,GlsProducto,IdNivel,IdUMCompra,IdUMVenta,IdMarca,AfectoIGV,IdMoneda,IdTipoProducto,IdFabricante," & _
                    "IdEmpresa,GlsObs,CodigoRapido,IdGrupo,IdTallaPeso,IndInsertaPrecioLista,CtaContable,IdGastosIngresos,CtaContable2_2010,EstProducto," & _
                    "CtaContable2,CtaContable_Relacionada)Values(" & _
                    "'" & CIdProducto & "','" & xl.Cells(NFil, 2).Value & "','" & TxtIdNivel.Text & "','" & TxtIdUM.Text & "','" & TxtIdUM.Text & "'," & _
                    "'" & TxtIdMarca.Text & "'," & Val("" & xl.Cells(NFil, 4).Value) & ",'" & TxtIdMoneda.Text & "','" & TxtIdTipoProducto.Text & "'," & _
                    "'" & xl.Cells(NFil, 1).Value & "','" & glsEmpresa & "','','','','',1,'" & TxtIdCtaContableVenta.Text & "','','','A'," & _
                    "'" & TxtIdCtaContableCompra.Text & "','" & TxtIdCtaContableRelacionada.Text & "')"
            
            Cn.Execute CSqlC
            
            CSqlC = "Insert Into ProductosAlmacen(IdAlmacen,IdProducto,Item,IdUMCompra,IdEmpresa,IdSucursal,CantidadStock,CostoUnit,Separacion,IdUbicacion)" & _
                    "Select A.IdAlmacen,'" & CIdProducto & "',@i:=@i+1,'" & TxtIdUM.Text & "',A.IdEmpresa,A.IdSucursal,0,0,0,'' " & _
                    "From (Select @i:=0) Z,Almacenes A " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "'"
            
            Cn.Execute CSqlC
            
            CSqlC = "Insert Into ProductosStock(IdEmpresa,IdSucursal,IdProducto,Stock,Separacion,Disponible)" & _
                    "Select A.IdEmpresa,A.IdSucursal,'" & CIdProducto & "',0,0,0 " & _
                    "From Sucursales A " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "'"
            
            Cn.Execute CSqlC
            
            CSqlC = "Insert Into Presentaciones(Item,IdProducto,IdUM,Factor,IdEmpresa)Values(" & _
                    "1,'" & CIdProducto & "','" & TxtIdUM.Text & "',1,'" & glsEmpresa & "')"

            Cn.Execute CSqlC

        End If
        
        NVVUnit = Val("" & xl.Cells(NFil, 3).Value)
        NIGVUnit = 0
        
        If Val("" & xl.Cells(NFil, 4).Value) = 1 Then
            NIGVUnit = NVVUnit * (glsIGV / 100)
        End If
        
        NPVUnit = NVVUnit + NIGVUnit
        
        CSqlC = "Select A.IdProducto " & _
                "From PreciosVenta A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdLista = '" & CIdLista & "' And A.IdProducto = '" & CIdProducto & "'"
        AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not RsC.EOF Then
            
            CSqlC = "Update PreciosVenta A " & _
                    "Set VVUnit = " & NVVUnit & ",IGVUnit = " & NIGVUnit & ",PVUnit = " & NPVUnit & " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdLista = '" & CIdLista & "' And A.IdProducto = '" & CIdProducto & "'"
            
        Else
        
            CSqlC = "Insert Into PreciosVenta(IdLista,IdProducto,IdUM,VVUnit,IGVUnit,PVUnit,IdEmpresa,CostoUnit,FactorUnit,Factor2Unit,MaxDcto)" & _
                    "Select '" & CIdLista & "',A.IdProducto,A.IdUMVenta," & NVVUnit & "," & NIGVUnit & "," & NPVUnit & ",A.IdEmpresa,0,0,0,0 " & _
                    "From Productos A " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & CIdProducto & "'"
        
        End If
        
        Cn.Execute CSqlC
        
        RsC.Close: Set RsC = Nothing
        
        NFil = NFil + 1
        
    Loop
    
    Cn.CommitTrans
    indTrans = False
    
    'Generar cabecera de automaticos
    
    MousePointer = 1
    Clipboard.Clear
    xl.ActiveWorkbook.Close False, False, False
    xl.Quit
    
    MsgBox "Fin de Proceso", vbInformation, App.Title
    
    Exit Sub
Err:
    If indTrans Then Cn.RollbackTrans
    MousePointer = 1
    Clipboard.Clear
    xl.ActiveWorkbook.Close False, False, False
    xl.Quit
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                     As String
    
    Me.top = 0
    Me.left = 0
    
    ValoresIniciales StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ValoresIniciales(StrMsgError As String)
On Error GoTo Err
Dim CSqlC                           As String
Dim RsC                             As New ADODB.Recordset

    CSqlC = "Select GlsParametro,ValParametro " & _
            "From Parametros " & _
            "Where IdEmpresa = '" & glsEmpresa & "' And GlsParametro In('IMP_PROD_NIVEL','IMP_PROD_TP','IMP_PROD_MARCA','IMP_PROD_MONEDA','IMP_PROD_UM'," & _
            "'IMP_PROD_CUENTAVENTA','IMP_PROD_CUENTACOMPRA','IMP_PROD_CUENTARELACIONADA')"
    AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
    If Not RsC.EOF Then
        RsC.Filter = "GlsParametro = 'IMP_PROD_NIVEL'"
        If Not RsC.EOF Then TxtIdNivel.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_TP'"
        If Not RsC.EOF Then TxtIdTipoProducto.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_MARCA'"
        If Not RsC.EOF Then TxtIdMarca.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_MONEDA'"
        If Not RsC.EOF Then TxtIdMoneda.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_UM'"
        If Not RsC.EOF Then TxtIdUM.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_CUENTAVENTA'"
        If Not RsC.EOF Then TxtIdCtaContableVenta.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_CUENTACOMPRA'"
        If Not RsC.EOF Then TxtIdCtaContableCompra.Text = Trim("" & RsC.Fields("ValParametro"))
        RsC.Filter = "GlsParametro = 'IMP_PROD_CUENTARELACIONADA'"
        If Not RsC.EOF Then TxtIdCtaContableRelacionada.Text = Trim("" & RsC.Fields("ValParametro"))
    End If
    RsC.Close: Set RsC = Nothing
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub TxtIdNivel_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsNivel.Text = Trim("" & traerCampo("Niveles A Inner Join TiposNiveles B On A.IdEmpresa = B.IdEmpresa And A.IdTipoNivel = B.IdTipoNivel", "A.GlsNivel", "A.IdNivel", TxtIdNivel.Text, False, "A.IdEmpresa = '" & glsEmpresa & "' And B.Peso = " & glsNumNiveles & ""))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdTipoProducto_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsTipoProducto.Text = Trim("" & traerCampo("Datos", "GlsDato", "IdDato", TxtIdTipoProducto.Text, False, "IdTipoDatos = '06'"))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdMarca_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsMarca.Text = Trim("" & traerCampo("Marcas", "GlsMarca", "IdMarca", TxtIdMarca.Text, True))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdMoneda_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsMoneda.Text = Trim("" & traerCampo("Monedas", "GlsMoneda", "IdMoneda", TxtIdMoneda.Text, False))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdUM_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsUM.Text = Trim("" & traerCampo("UnidadMedida", "GlsUM", "IdUM", TxtIdUM.Text, False))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdCtaContableVenta_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsCtaContableVenta.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtIdCtaContableVenta.Text, True, "GradoCuenta = " & LeeParametroConta("GRADO_MAXIMO") & ""))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdCtaContableCompra_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsCtaContableCompra.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtIdCtaContableCompra.Text, True, "GradoCuenta = " & LeeParametroConta("GRADO_MAXIMO") & ""))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdCtaContableRelacionada_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    TxtGlsCtaContableRelacionada.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtIdCtaContableRelacionada.Text, True, "GradoCuenta = " & LeeParametroConta("GRADO_MAXIMO") & ""))
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
