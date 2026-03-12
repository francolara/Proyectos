VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmActualizaMaxDcto 
   Caption         =   "Actualización del Máximo Descuento"
   ClientHeight    =   2625
   ClientLeft      =   3135
   ClientTop       =   3060
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   8910
   Begin VB.CommandButton BtnSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1980
      Width           =   1245
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1980
      Width           =   1245
   End
   Begin VB.Frame fraContenido 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   1770
      Left            =   75
      TabIndex        =   9
      Top             =   0
      Width           =   8790
      Begin VB.CommandButton cmbAyudaMarca 
         Height          =   315
         Left            =   8145
         Picture         =   "FrmActualizaMaxDcto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   690
         Width           =   390
      End
      Begin VB.Frame fraNivel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   45
         TabIndex        =   10
         Top             =   270
         Width           =   8625
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   4
            Left            =   8100
            Picture         =   "FrmActualizaMaxDcto.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   3
            Left            =   8100
            Picture         =   "FrmActualizaMaxDcto.frx":0714
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1080
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   2
            Left            =   8100
            Picture         =   "FrmActualizaMaxDcto.frx":0A9E
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   720
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   1
            Left            =   8100
            Picture         =   "FrmActualizaMaxDcto.frx":0E28
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   0
            Left            =   8100
            Picture         =   "FrmActualizaMaxDcto.frx":11B2
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   0
            Left            =   1305
            TabIndex        =   0
            Tag             =   "TidNivelPred"
            Top             =   45
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            MaxLength       =   8
            Container       =   "FrmActualizaMaxDcto.frx":153C
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   16
            Top             =   30
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
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
            Container       =   "FrmActualizaMaxDcto.frx":1558
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   1
            Left            =   1305
            TabIndex        =   1
            Tag             =   "TidNivelPred"
            Top             =   390
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            MaxLength       =   8
            Container       =   "FrmActualizaMaxDcto.frx":1574
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   17
            Top             =   390
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
            BackColor       =   12648447
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
            Container       =   "FrmActualizaMaxDcto.frx":1590
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   2
            Left            =   1305
            TabIndex        =   2
            Tag             =   "TidNivelPred"
            Top             =   750
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            MaxLength       =   8
            Container       =   "FrmActualizaMaxDcto.frx":15AC
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   2
            Left            =   2280
            TabIndex        =   18
            Top             =   750
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
            BackColor       =   12648447
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
            Container       =   "FrmActualizaMaxDcto.frx":15C8
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   3
            Left            =   1305
            TabIndex        =   3
            Tag             =   "TidNivelPred"
            Top             =   1110
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            MaxLength       =   8
            Container       =   "FrmActualizaMaxDcto.frx":15E4
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   3
            Left            =   2280
            TabIndex        =   19
            Top             =   1110
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
            BackColor       =   12648447
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
            Container       =   "FrmActualizaMaxDcto.frx":1600
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   285
            Index           =   4
            Left            =   1305
            TabIndex        =   4
            Tag             =   "TidNivelPred"
            Top             =   1470
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            MaxLength       =   8
            Container       =   "FrmActualizaMaxDcto.frx":161C
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   285
            Index           =   4
            Left            =   2280
            TabIndex        =   20
            Top             =   1470
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   503
            BackColor       =   12648447
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
            Container       =   "FrmActualizaMaxDcto.frx":1638
            Vacio           =   -1  'True
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   25
            Top             =   1485
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   24
            Top             =   1125
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   23
            Top             =   765
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   22
            Top             =   405
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   45
            Width           =   405
         End
      End
      Begin CATControls.CATTextBox TxtDctoListaPrec 
         Height          =   315
         Left            =   1350
         TabIndex        =   6
         Top             =   1215
         Width           =   465
         _ExtentX        =   820
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "FrmActualizaMaxDcto.frx":1654
         Text            =   "-------  0  -------"
         Estilo          =   3
         TextoInicio     =   "0"
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Marca 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   735
         Width           =   915
         _ExtentX        =   1614
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
         MaxLength       =   8
         Container       =   "FrmActualizaMaxDcto.frx":1670
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Marca 
         Height          =   315
         Left            =   2310
         TabIndex        =   28
         Top             =   735
         Width           =   5805
         _ExtentX        =   10239
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
         Container       =   "FrmActualizaMaxDcto.frx":168C
         Vacio           =   -1  'True
      End
      Begin VB.Label LblMarca 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   29
         Top             =   780
         Width           =   495
      End
      Begin VB.Label LblMaxDcto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Max % Dcto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   26
         Top             =   1215
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmActualizaMaxDcto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
    
    Unload Me
    
End Sub

Private Sub CmbAyudaMarca_Click()
Dim StrMsgError                         As String
On Error GoTo Err
    
    mostrarAyuda "MARCA", txtCod_Marca, txtGls_Marca
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmbAyudaNivel_Click(Index As Integer)
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

Private Sub cmdaceptar_Click()
Dim StrMsgError                         As String
Dim cWhereNiveles                       As String
On Error GoTo Err
    
    cWhereNiveles = ""
    
    If Len(Trim(txtCod_Nivel(0).Text)) > 0 Then
        cWhereNiveles = cWhereNiveles & "And N.IdNivel" & Format(glsNumNiveles, "00") & " = '" & txtCod_Nivel(0).Text & "' "
        If Len(Trim(txtCod_Nivel(1).Text)) > 0 Then
            cWhereNiveles = cWhereNiveles & "And N.IdNivel" & Format(glsNumNiveles - 1, "00") & " = '" & txtCod_Nivel(1).Text & "' "
            If Len(Trim(txtCod_Nivel(2).Text)) > 0 Then
                cWhereNiveles = cWhereNiveles & "And N.IdNivel" & Format(glsNumNiveles - 2, "00") & " = '" & txtCod_Nivel(2).Text & "' "
            End If
        End If
    End If
    
    CSqlC = "Update Vw_Niveles N " & _
            "Inner Join Productos A " & _
                "On N.IdEmpresa = A.IdEmpresa And N.IdNivel01 = A.IdNivel " & _
            "Inner Join PreciosVenta B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
            "Set B.MaxDcto = " & Val("" & TxtDctoListaPrec.Text) & " " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdMarca Like'%" & txtCod_Marca.Text & "%' " & cWhereNiveles
    
    Cn.Execute (CSqlC)
        
    MsgBox "Fin de Proceso", vbInformation, App.Title
        
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txtGls_Nivel(0).Text = "TODO LOS GRUPOS"
    txtGls_Nivel(1).Text = "TODA LAS CATEGORIAS"
    txtGls_Nivel(2).Text = "TODA LAS SUB CATEGORIAS"
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
Dim rsj             As New ADODB.Recordset
On Error GoTo Err

    'jalamos tipos nivel
    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    
    Me.Height = Me.Height + (355 * glsNumNiveles)
    
    fraContenido.Height = fraContenido.Height + (355 * glsNumNiveles)
    
    fraNivel.Height = 355 * glsNumNiveles
    
    LblMarca.top = LblMarca.top + (355 * glsNumNiveles)
    txtCod_Marca.top = txtCod_Marca.top + (355 * glsNumNiveles)
    txtGls_Marca.top = txtGls_Marca.top + (355 * glsNumNiveles)
    CmbAyudaMarca.top = CmbAyudaMarca.top + (355 * glsNumNiveles)
    LblMaxDcto.top = LblMaxDcto.top + (355 * glsNumNiveles)
    TxtDctoListaPrec.top = TxtDctoListaPrec.top + (355 * glsNumNiveles)
    CmdAceptar.top = CmdAceptar.top + (355 * glsNumNiveles)
    BtnSalir.top = BtnSalir.top + (355 * glsNumNiveles)
    
    i = 0
    
    Do While Not rsj.EOF
        
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        
        rsj.MoveNext
        
        i = i + 1
    
    Loop
    
    If rsj.State = 1 Then rsj.Close
    Set rsj = Nothing
    
    Exit Sub
    
Err:
If rsj.State = 1 Then rsj.Close
Set rsj = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Marca_Change()
    
    txtGls_Marca.Text = traerCampo("marcas", "GlsMarca", "idMarca", txtCod_Marca.Text, True)
    
End Sub

Private Sub txtCod_Nivel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 8 Then
        If txtCod_Nivel(0).Text <> "" Then
                txtCod_Nivel(0).Text = ""
                txtGls_Nivel(0).Text = "TODO LOS GRUPOS"
                
                txtCod_Nivel(1).Text = ""
                txtGls_Nivel(1).Text = "TODA LAS CATEGORIAS"
                
                txtCod_Nivel(2).Text = ""
                txtGls_Nivel(2).Text = "TODA LAS SUB CATEGORIAS"
            Exit Sub
        End If
    End If
    
End Sub
