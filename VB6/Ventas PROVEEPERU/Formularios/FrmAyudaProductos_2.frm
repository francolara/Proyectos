VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmAyudaProductos_2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Productos"
   ClientHeight    =   8895
   ClientLeft      =   795
   ClientTop       =   855
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAyudaProductos_2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14670
   Begin VB.CommandButton cmbProdOtrasSucursales 
      Caption         =   "Consultar en otras sucursales"
      Height          =   450
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8370
      Width           =   2475
   End
   Begin VB.Frame fraContenido 
      Appearance      =   0  'Flat
      Caption         =   "Filtros:"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   45
      TabIndex        =   10
      Top             =   60
      Width           =   11445
      Begin VB.Frame fraNivel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   360
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Width           =   8625
         Begin VB.CommandButton cmbAyudaNivel 
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
            Index           =   0
            Left            =   8100
            Picture         =   "FrmAyudaProductos_2.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   45
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
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
            Index           =   1
            Left            =   8100
            Picture         =   "FrmAyudaProductos_2.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
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
            Index           =   2
            Left            =   8100
            Picture         =   "FrmAyudaProductos_2.frx":0720
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   720
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
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
            Index           =   3
            Left            =   8100
            Picture         =   "FrmAyudaProductos_2.frx":0AAA
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1080
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
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
            Index           =   4
            Left            =   8100
            Picture         =   "FrmAyudaProductos_2.frx":0E34
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1440
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   0
            Left            =   1305
            TabIndex        =   19
            Tag             =   "TidNivelPred"
            Top             =   30
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
            Container       =   "FrmAyudaProductos_2.frx":11BE
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   20
            Top             =   30
            Width           =   5790
            _ExtentX        =   10213
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
            Container       =   "FrmAyudaProductos_2.frx":11DA
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   1
            Left            =   1305
            TabIndex        =   21
            Tag             =   "TidNivelPred"
            Top             =   390
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
            Container       =   "FrmAyudaProductos_2.frx":11F6
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   22
            Top             =   390
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
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
            Container       =   "FrmAyudaProductos_2.frx":1212
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   2
            Left            =   1305
            TabIndex        =   23
            Tag             =   "TidNivelPred"
            Top             =   750
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
            Container       =   "FrmAyudaProductos_2.frx":122E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   2
            Left            =   2280
            TabIndex        =   24
            Top             =   750
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
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
            Container       =   "FrmAyudaProductos_2.frx":124A
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   3
            Left            =   1305
            TabIndex        =   25
            Tag             =   "TidNivelPred"
            Top             =   1110
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
            Container       =   "FrmAyudaProductos_2.frx":1266
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   3
            Left            =   2280
            TabIndex        =   26
            Top             =   1110
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
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
            Container       =   "FrmAyudaProductos_2.frx":1282
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   4
            Left            =   1305
            TabIndex        =   27
            Tag             =   "TidNivelPred"
            Top             =   1470
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
            Container       =   "FrmAyudaProductos_2.frx":129E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   4
            Left            =   2280
            TabIndex        =   28
            Top             =   1470
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
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
            Container       =   "FrmAyudaProductos_2.frx":12BA
            Vacio           =   -1  'True
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   33
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   135
            TabIndex        =   32
            Top             =   405
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   135
            TabIndex        =   31
            Top             =   765
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   135
            TabIndex        =   30
            Top             =   1125
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   135
            TabIndex        =   29
            Top             =   1485
            Width           =   345
         End
      End
      Begin CATControls.CATTextBox TxtBusq 
         Height          =   315
         Left            =   1365
         TabIndex        =   0
         Top             =   1140
         Width           =   6750
         _ExtentX        =   11906
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
         Container       =   "FrmAyudaProductos_2.frx":12D6
         Vacio           =   -1  'True
      End
      Begin VB.Label lblBusq 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   195
         TabIndex        =   11
         Top             =   1200
         Width           =   645
      End
   End
   Begin VB.Frame fraPresentaciones 
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
      Height          =   1605
      Left            =   90
      TabIndex        =   9
      Top             =   6705
      Width           =   14565
      Begin DXDBGRIDLibCtl.dxDBGrid gPresentaciones 
         Height          =   1305
         Left            =   60
         OleObjectBlob   =   "FrmAyudaProductos_2.frx":12F2
         TabIndex        =   34
         Top             =   180
         Width           =   14475
      End
   End
   Begin VB.Frame fraTipoProd 
      Appearance      =   0  'Flat
      Caption         =   " Tipo "
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   11580
      TabIndex        =   5
      Top             =   60
      Width           =   3045
      Begin VB.OptionButton opt_MateriaPrima 
         Caption         =   "Materia Prima"
         Height          =   240
         Left            =   675
         TabIndex        =   8
         Top             =   1335
         Width           =   1290
      End
      Begin VB.OptionButton opt_Servicios 
         Caption         =   "Servicios"
         Height          =   240
         Left            =   675
         TabIndex        =   7
         Top             =   855
         Width           =   1065
      End
      Begin VB.OptionButton opt_Producto 
         Caption         =   "Productos"
         Height          =   240
         Left            =   675
         TabIndex        =   6
         Top             =   375
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame fraGrilla 
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
      Height          =   4365
      Left            =   75
      TabIndex        =   3
      Top             =   1920
      Width           =   14565
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   4125
         Left            =   90
         OleObjectBlob   =   "FrmAyudaProductos_2.frx":3C58
         TabIndex        =   4
         Top             =   180
         Width           =   14445
      End
   End
   Begin VB.Label lblPresentaciones 
      Appearance      =   0  'Flat
      Caption         =   "Otras Presentaciones:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      TabIndex        =   12
      Top             =   6360
      Width           =   3435
   End
   Begin VB.Label LblReg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "(0) Registros"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12735
      TabIndex        =   2
      Top             =   6345
      Width           =   1905
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Presionar Enter en el registro para obtener el resultado "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10575
      TabIndex        =   1
      Top             =   8415
      Width           =   4005
   End
End
Attribute VB_Name = "FrmAyudaProductos_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SqlAdic As String
Private sqlBus As String
Private sqlCond As String
Private SRptBus(3) As String
Private EsNuevo As Boolean
Private indAlmacen As Boolean

Private indValidaStock As Boolean
Private indPedido As Boolean

Private strCodAlmacen As String
Private indUMVenta As Boolean
Private indMostrarPresentaciones As Boolean

Private strCodLista As String

Private indMovNivel As Boolean
Private intFoco As Integer '0 = Texto,1 = Grilla productos, 2 = Grilla presentaciones

Private Sub CmdBusq_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub cmbAyudaNivel_Click(Index As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    peso = Index + 1
    strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
    strCondPred = ""
    If peso > 1 Then
        strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
    End If
    mostrarAyuda "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    
End Sub

Private Sub cmbProdOtrasSucursales_Click()
On Error GoTo Err
Dim strCodProd As String, StrMsgError As String

    strCodProd = g.Columns.ColumnByFieldName("idProducto").Value
    frmProdOtrasSucursales.MostrarForm strCodProd, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Deactivate()
    
    SqlAdic = ""
    If EsNuevo = False Then
        TxtBusq.Text = ""
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
            Case 0
                g.SetFocus
            Case 1
                gPresentaciones.SetFocus
            Case 2
                TxtBusq.SetFocus
            End Select
    End Select

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    Me.Caption = "Ayuda de productos"
    ConfGrid g, False, False, False, False
    ConfGrid gPresentaciones, False, False, False, False
    
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    EsNuevo = True
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Unload(Cancel As Integer)

    SqlAdic = ""
    TxtBusq.Text = ""

End Sub

Private Sub g_GotFocus()

    intFoco = 1

End Sub

Private Sub G_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError As String
    
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub g_OnDblClick()
    
    g_OnKeyDown 13, 1

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 13:
            SRptBus(0) = g.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = g.Columns.ColumnByFieldName("GlsProducto").Value '+ Space(150) + Trim(CStr(g.Columns.ColumnByFieldName("cod").Value))
            SRptBus(2) = g.Columns.ColumnByFieldName("idUMVenta").Value
            
            g.Dataset.Close
            g.Dataset.Active = False
            Me.Hide
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
                Case 0
                    g.SetFocus
                Case 1
                    gPresentaciones.SetFocus
                Case 2
                    TxtBusq.SetFocus
            End Select
    End Select

End Sub

Private Sub gPresentaciones_GotFocus()
    
    intFoco = 2

End Sub

Private Sub gPresentaciones_OnDblClick()
    
    gPresentaciones_OnKeyDown 13, 1

End Sub

Private Sub gPresentaciones_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 13:
            SRptBus(0) = g.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = g.Columns.ColumnByFieldName("GlsProducto").Value '+ Space(150) + Trim(CStr(g.Columns.ColumnByFieldName("cod").Value))
            SRptBus(2) = gPresentaciones.Columns.ColumnByFieldName("idUM").Value
            
            g.Dataset.Close
            g.Dataset.Active = False
            Me.Hide
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
                Case 0
                    g.SetFocus
                Case 1
                    gPresentaciones.SetFocus
                Case 2
                    TxtBusq.SetFocus
            End Select
    End Select

End Sub

Private Sub opt_MateriaPrima_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub opt_Producto_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub opt_Servicios_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub TxtBusq_Change()
Dim StrMsgError As String

    If EsNuevo = False Then
        If glsEnterAyudaProductos = False Then
            fill StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub TxtBusq_GotFocus()
    
    intFoco = 0

End Sub

Private Sub TxtBusq_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyDown Then g.SetFocus
    If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1

End Sub

Private Sub TxtBusq_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        fill StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If g.Count > 1 Then g.SetFocus
    End If
    EsNuevo = False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub fill(ByRef StrMsgError As String)
On Error GoTo Err

    sqlBus = setSqlAlm(strCodAlmacen)
    sqlCond = sqlBus + " like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') "
    
    If txtCod_Nivel(glsNumNiveles - 1).Text <> "" Then
        sqlCond = sqlCond & " AND idNivel = '" & txtCod_Nivel(glsNumNiveles - 1).Text & "'"
    End If
    
    If opt_Producto.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06001'"
    ElseIf opt_MateriaPrima.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06003'"
    End If
    sqlCond = sqlCond & " AND estProducto = 'A' "
    sqlCond = sqlCond & SqlAdic & " order by 1"
    
    With g
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sqlCond
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idProducto"
    End With
    
    LblReg.Caption = "(" + Format(g.Count, "0") + ")Registros"
    
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Public Sub Execute(ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
'    Dim strMsgError As String
'    Dim intI As Integer
'    pblnAceptar = False
'    MousePointer = 0
'
'    SRptBus(0) = ""
'
'    SqlAdic = strParAdic
'
'    sqlBus = setSql(strParAyuda)
'
'    fill
'
'    Me.Show vbModal
'    If SRptBus(0) <> "" Then
'        TextBox1.Text = SRptBus(0)
'        TextBox2.Text = SRptBus(1)
'    End If
End Sub

Private Function setSqlAlm(strAlm As String) As String
Dim strCampoUM As String
Dim strStockUM As String
Dim strCantidad As String
Dim strTablaPresentaciones As String

    If glsVisualizaCodFab = "N" Then
        g.Columns.ColumnByFieldName("IdFabricante").Visible = False
    End If

    strCampoUM = "idUMVenta"
    strStockUM = "CantidadStockUV"
    strCantidad = "(a.CantidadStock / f.Factor )" 'Es la cantidad de venta
    
    strTablaPresentaciones = " INNER JOIN presentaciones f ON p.idEmpresa = f.idEmpresa AND p.idProducto = f.idProducto AND p." & strCampoUM & " = f.idUM "
        
    If indUMVenta = False Then
        strCampoUM = "idUMCompra"
        strStockUM = "CantidadStockUC"
        strCantidad = "a.CantidadStock" 'Es la cantidad de compra
        strTablaPresentaciones = ""
    End If
    
    g.Columns.ColumnByFieldName("Stock").Visible = False
    If opt_Servicios.Value = False Then
        If indPedido = False Then
            setSqlAlm = "SELECT p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(" & strCantidad & ",2) as Stock, t.GlsTallaPeso, p.idfabricante " & _
                        "FROM productos p " & _
                        "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                        "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                        "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                        "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND a.idSucursal = '" & glsSucursal & "' " & _
                                                      "AND a.idAlmacen = '" & strAlm & "' AND p.idProducto = a.idProducto " & _
                                                      "AND p." & strCampoUM & " = a.idUMCompra " & strTablaPresentaciones & _
                        "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                        "WHERE p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')" & _
                        "AND p.idEmpresa = '" & glsEmpresa & "' "
    
            If indValidaStock Then
                 setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
            End If
            g.Columns.ColumnByFieldName("Stock").Visible = True
            setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
        
        Else
            setSqlAlm = "SELECT p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(0,2) as Stock, t.GlsTallaPeso " & _
                        "FROM productos p " & _
                        "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                        "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                        "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                        "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND p.idProducto = a.idProducto AND a.idAlmacen  = '" & strAlm & "' " & _
                        "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                        "WHERE p.idEmpresa = '" & glsEmpresa & "' AND (p.GlsProducto "
        End If
        
    Else
        setSqlAlm = "SELECT p.idProducto,p.GlsProducto,'' as GlsMarca, p.idUMVenta, u.GlsUM, o.GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(0,2) as Stock, '' AS GlsTallaPeso " & _
                    "FROM productos p,monedas o, unidadMedida u " & _
                    "WHERE p.idMoneda  = o.idMoneda AND p.idEmpresa = '" & glsEmpresa & "' AND p.idTipoProducto = '06002' AND p.idUMVenta = u.idUM AND (p.GlsProducto "
    End If
    
End Function

Public Sub ExecuteReturnTextAlm(ByVal strAlm As String, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal ValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
On Error GoTo Err
    
    MousePointer = 0
    
    'Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    
    'Pasamos valores de parametros a las variables privadas a nivel de form
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    strCodAlmacen = strAlm
    SqlAdic = strParAdic
    indUMVenta = indVarUMVenta
    indMostrarPresentaciones = indVarMostrarPresentaciones

    'Asignamos valores
    fraTipoProd.Visible = indMostrarTP
    
    If indMostrarPresentaciones = False Then
        Me.Height = fraPresentaciones.top + 350
        lblPresentaciones.Visible = False
    End If
    
    Select Case TipoProd
    Case 1 'productos
        opt_Producto.Value = True
    Case 2 'servicios
        opt_Servicios.Value = True
    Case 3 'materia prima
        opt_MateriaPrima.Value = True
    End Select
    
    'Filtramos
    fill StrMsgError
    If StrMsgError <> "" Then StrMsgError = Err.Description
    
    Me.Show vbModal
    
    'Devolvemos valores
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
        strCodUM = SRptBus(2)
    End If
    
    Set g.DataSource = Nothing
    Set gPresentaciones.DataSource = Nothing
    
    Unload Me
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Public Sub ExecuteKeyasciiReturnTextAlm(ByVal KeyAscii As Integer, strAlm As String, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal ValidaStock As String, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
On Error GoTo Err

    MousePointer = 0
    
    'Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    
    'Pasamos valores de parametros a las variables privadas a nivel de form
    strCodAlmacen = strAlm
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    SqlAdic = strParAdic
    indUMVenta = indVarUMVenta
    indMostrarPresentaciones = indVarMostrarPresentaciones
    
    'Asignamos valores
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    fraTipoProd.Visible = indMostrarTP
    
    If indMostrarPresentaciones = False Then
        Me.Height = fraPresentaciones.top + 350
        lblPresentaciones.Visible = False
    End If
    
    Select Case TipoProd
    Case 1 'productos
        opt_Producto.Value = True
    Case 2 'servicios
        opt_Servicios.Value = True
    Case 3 'materia prima
        opt_MateriaPrima.Value = True
    End Select
    
    'Filtramos
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Show vbModal
    
    'Devolvemos valores
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
        strCodUM = SRptBus(2)
    End If

    Set g.DataSource = Nothing
    Set gPresentaciones.DataSource = Nothing
    
    Unload Me

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub listaOtrasPresentaciones(ByRef StrMsgError As String)
On Error GoTo Err

    If indMostrarPresentaciones = False Then Exit Sub
    
    csql = "SELECT p.idUM,u.abreUM as GlsUM,Format(r.factor,2) AS factor,p.VVUnit AS VVUnit,p.IGVUnit AS IGVUnit,p.PVUnit AS PVUnit " & _
            "FROM preciosventa p,unidadMedida u, presentaciones r " & _
            "WHERE p.idUM = u.idUM " & _
            "AND p.idProducto = '" & g.Columns.ColumnByFieldName("idProducto").Value & "' " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idUM = r.idUM " & _
            "AND p.idProducto = r.idProducto " & _
            "AND r.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idLista = '" & strCodLista & "' ORDER BY r.factor ASC"
               
    With gPresentaciones
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idUM"
    End With
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsj As New ADODB.Recordset
Dim i As Integer

    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    fraNivel.Height = 355 * glsNumNiveles
    i = 0
    Do While Not rsj.EOF
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    
    TxtBusq.top = fraNivel.top + fraNivel.Height + 35
    lblBusq.top = TxtBusq.top
    fraContenido.Height = TxtBusq.top + TxtBusq.Height + 100
    If fraTipoProd.Height > fraContenido.Height Then
        fraContenido.Height = fraTipoProd.Height
    Else
        fraTipoProd.Height = fraContenido.Height
        fraGrilla.top = fraTipoProd.top + fraTipoProd.Height
        fraGrilla.Height = fraPresentaciones.top - (fraGrilla.top + lblPresentaciones.Height)
        g.Height = fraGrilla.Height - 200
    End If
    
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    
    Exit Sub

Err:
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub txtCod_Nivel_Change(Index As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If indMovNivel Then Exit Sub
    
    txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)
    indMovNivel = True
    For i = Index + 1 To txtCod_Nivel.Count - 1
        txtCod_Nivel(i).Text = ""
        txtGls_Nivel(i).Text = ""
    Next
    indMovNivel = False
    
    If glsNumNiveles = Index + 1 Then
        fill StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Nivel_KeyPress(Index As Integer, KeyAscii As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    If KeyAscii <> 13 Then
        peso = Index + 1
        strCodJerarquia = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
        strCondPred = ""
        If peso > 1 Then
            strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
        End If
        mostrarAyudaKeyascii KeyAscii, "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    End If

End Sub
