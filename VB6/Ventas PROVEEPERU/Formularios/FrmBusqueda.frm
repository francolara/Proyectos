VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmBusqueda 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   6510
   ClientLeft      =   10860
   ClientTop       =   4515
   ClientWidth     =   8280
   Icon            =   "FrmBusqueda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8280
   Begin VB.Frame FraRegVehiculo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   60
      TabIndex        =   27
      Top             =   1470
      Visible         =   0   'False
      Width           =   8145
      Begin VB.Frame Frame4 
         Height          =   1905
         Left            =   1290
         TabIndex        =   29
         Top             =   570
         Width           =   5625
         Begin VB.ComboBox CbxMarcas 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1170
            Width           =   4080
         End
         Begin CATControls.CATTextBox txtGls_Vehiculo 
            Height          =   315
            Left            =   1365
            TabIndex        =   33
            Tag             =   "TglsVehiculo"
            Top             =   360
            Width           =   4050
            _ExtentX        =   7144
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
            Container       =   "FrmBusqueda.frx":000C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Placa 
            Height          =   315
            Left            =   1365
            TabIndex        =   34
            Tag             =   "TGlsPlaca"
            Top             =   750
            Width           =   1500
            _ExtentX        =   2646
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
            MaxLength       =   16
            Container       =   "FrmBusqueda.frx":0028
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label4 
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
            Left            =   180
            TabIndex        =   37
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Marca"
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
            Left            =   180
            TabIndex        =   36
            Top             =   1215
            Width           =   450
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Placa"
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
            Left            =   180
            TabIndex        =   35
            Top             =   810
            Width           =   390
         End
      End
      Begin VB.CommandButton cmdGrabarVehiculo 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   2970
         TabIndex        =   0
         Top             =   2610
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelaVehiculo 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4230
         TabIndex        =   28
         Top             =   2610
         Width           =   1095
      End
      Begin CATControls.CATTextBox CATTextBox6 
         Height          =   315
         Left            =   1650
         TabIndex        =   30
         Top             =   7275
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
         Container       =   "FrmBusqueda.frx":0044
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REGISTRO RAPIDO DE VEHICULOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   30
         TabIndex        =   32
         Top             =   120
         Width           =   8100
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº Brevete"
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
         Left            =   135
         TabIndex        =   31
         Top             =   7350
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdopera 
      Caption         =   "&REGISTRO RAPIDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3390
      TabIndex        =   26
      Top             =   6000
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Frame FraRegChofer 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   60
      TabIndex        =   9
      Top             =   1470
      Visible         =   0   'False
      Width           =   8145
      Begin VB.CommandButton cmdopera 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4110
         TabIndex        =   22
         Top             =   2610
         Width           =   1095
      End
      Begin VB.CommandButton cmdopera 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   21
         Top             =   2610
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   1905
         Left            =   240
         TabIndex        =   10
         Top             =   570
         Width           =   7665
         Begin CATControls.CATTextBox txtGls_Paterno 
            Height          =   315
            Left            =   1335
            TabIndex        =   11
            Tag             =   "TapellidoPaterno"
            Top             =   300
            Width           =   6240
            _ExtentX        =   11007
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
            MaxLength       =   80
            Container       =   "FrmBusqueda.frx":0060
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Materno 
            Height          =   315
            Left            =   1335
            TabIndex        =   12
            Tag             =   "TapellidoMaterno"
            Top             =   645
            Width           =   6240
            _ExtentX        =   11007
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
            MaxLength       =   80
            Container       =   "FrmBusqueda.frx":007C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nombres 
            Height          =   315
            Left            =   1335
            TabIndex        =   13
            Tag             =   "Tnombres"
            Top             =   1005
            Width           =   6240
            _ExtentX        =   11007
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
            MaxLength       =   80
            Container       =   "FrmBusqueda.frx":0098
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_RUC 
            Height          =   315
            Left            =   1335
            TabIndex        =   14
            Tag             =   "Truc"
            Top             =   1380
            Width           =   2115
            _ExtentX        =   3731
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
            Container       =   "FrmBusqueda.frx":00B4
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox TxtBrevete 
            Height          =   315
            Left            =   5475
            TabIndex        =   15
            Tag             =   "Truc"
            Top             =   1380
            Width           =   2085
            _ExtentX        =   3678
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
            MaxLength       =   20
            Container       =   "FrmBusqueda.frx":00D0
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "DNI"
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
            Left            =   75
            TabIndex        =   20
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label lblApePaterno 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Apellido Paterno"
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
            Left            =   60
            TabIndex        =   19
            Top             =   390
            Width           =   1170
         End
         Begin VB.Label lblApeMaterno 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Apellido Materno"
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
            Left            =   60
            TabIndex        =   18
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblNombres 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nombres"
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
            Left            =   60
            TabIndex        =   17
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Brevete"
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
            Left            =   4125
            TabIndex        =   16
            Top             =   1440
            Width           =   570
         End
      End
      Begin CATControls.CATTextBox CATTextBox11 
         Height          =   315
         Left            =   1650
         TabIndex        =   23
         Top             =   7275
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
         Container       =   "FrmBusqueda.frx":00EC
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº Brevete"
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
         Left            =   135
         TabIndex        =   25
         Top             =   7350
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REGISTRO RAPIDO DE CHOFERES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   8100
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   75
      TabIndex        =   7
      Top             =   540
      Width           =   8145
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   5115
         Left            =   75
         OleObjectBlob   =   "FrmBusqueda.frx":0108
         TabIndex        =   2
         Top             =   150
         Width           =   8010
      End
   End
   Begin VB.CommandButton Cmdbusq 
      DownPicture     =   "FrmBusqueda.frx":17EE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7650
      Picture         =   "FrmBusqueda.frx":1B99
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox TxtBusq 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   945
      MaxLength       =   255
      TabIndex        =   1
      Top             =   165
      Width           =   7245
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      MaxLength       =   255
      TabIndex        =   8
      Top             =   4575
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Presionar Enter en el registro para obtener el resultado "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   90
      TabIndex        =   4
      Top             =   5985
      Width           =   3120
   End
   Begin VB.Label LblReg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "(0) Registros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6300
      TabIndex        =   5
      Top             =   5985
      Width           =   1905
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   135
      TabIndex        =   3
      Top             =   225
      Width           =   735
   End
End
Attribute VB_Name = "FrmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Dim cadRelacion As String
Dim cadCondicion As String
Private SqlAdic As String
Private sqlBus As String
Private sqlCond As String
Private SRptBus(2) As String
Private strAyuda As String
Private EsNuevo As Boolean


Private Sub CmdBusq_Click()
    
    sqlBus = setSql(strAyuda)
    fill

End Sub

Private Sub CmdCancelaVehiculo_Click()
    FraRegVehiculo.Visible = False
End Sub

Private Sub cmdGrabarVehiculo_Click()
On Error GoTo Err
Dim strCodigo       As String
Dim strMsg          As String
Dim CPrefijo        As String
Dim StrCodPersona   As String
Dim StrCadSql       As String
Dim StrMsgError     As String

        If Trim("" & txtGls_Vehiculo.Text) = "" Then
            StrMsgError = "Faltan datos del vechiculo, verifique.": GoTo Err
        End If
        If Trim("" & txtGls_Placa.Text) = "" Then
            StrMsgError = "Faltan datos del vechiculo, verifique.": GoTo Err
        End If
        If Trim("" & right(CbxMarcas.Text, 8)) = "" Then
            StrMsgError = "Faltan datos del vechiculo, verifique.": GoTo Err
        End If
        
        validaHomonimia "Vehiculos", "GlsVehiculo", "idVehiculo", txtGls_Vehiculo.Text, "", True, StrMsgError, " or GlsPlaca = '" & txtGls_Placa.Text & "'"
        If StrMsgError <> "" Then GoTo Err
    
        StrCodPersona = GeneraCorrelativoAnoMes("Vehiculos", "idVehiculo")
        
        StrCadSql = "EXEC spu_RegVehiculos '" & StrCodPersona & "','" & glsEmpresa & "','" & Trim("" & txtGls_Vehiculo.Text) & "'," & _
                    "'" & Trim("" & txtGls_Placa.Text) & "','" & Trim("" & right(CbxMarcas.Text, 8)) & "'"
        Cn.Execute StrCadSql
        
        Call fill

        FraRegVehiculo.Visible = False
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdopera_Click(Index As Integer)
On Error GoTo Err
Dim strCodigo       As String
Dim strMsg          As String
Dim CPrefijo        As String
Dim StrCodPersona   As String
Dim StrCadSql       As String
Dim StrMsgError     As String
Dim RsVarios        As String

    If Index = 2 Then
        If strAyuda = "CHOFER" Then
            FraRegChofer.Visible = True
            txtGls_Paterno.Text = ""
            txtGls_Materno.Text = ""
            txtGls_Nombres.Text = ""
            txt_RUC.Text = ""
            TxtBrevete.Text = ""
        End If
        
        If strAyuda = "VEHICULO" Then
            FraRegVehiculo.Visible = True
            txtGls_Vehiculo.Text = ""
            txtGls_Placa.Text = ""
            
            llenaCombo CbxMarcas, "datos", "GlsDato", False, "idDato", "idTipoDatos = '07'"
            
        End If
        
        Exit Sub
    End If
    
    If Index = 0 Then
        
        If Trim("" & txtGls_Paterno.Text) = "" Then
            StrMsgError = "Faltan datos del chofer, verifique.": GoTo Err
        End If
        If Trim("" & txtGls_Materno.Text) = "" Then
            StrMsgError = "Faltan datos del chofer, verifique.": GoTo Err
        End If
        If Trim("" & txtGls_Nombres.Text) = "" Then
            StrMsgError = "Faltan datos del chofer, verifique.": GoTo Err
        End If
        If Trim("" & txt_RUC.Text) = "" And Len(Trim("" & txt_RUC.Text)) <> 8 Then
             StrMsgError = "Ingrese un DNI valido, verifique.": GoTo Err
        End If
        If Trim("" & TxtBrevete.Text) = "" Then
            StrMsgError = "Faltan datos del chofer, verifique.": GoTo Err
        End If
        
        If Trim(txt_RUC.Text) <> "" Then
            validaHomonimia "personas", "ruc", "idPersona", txt_RUC.Text, "", False, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    
        CPrefijo = Year(getFechaSistema) & Format(Month(getFechaSistema), "00")
        StrCodPersona = GeneraCorrelativoAnoMesNuevo("personas", "idPersona", CPrefijo, False)
        
        StrCadSql = "EXEC spu_RegChoferes '" & StrCodPersona & "','" & glsEmpresa & "','" & Trim("" & txtGls_Paterno.Text) & "'," & _
                    "'" & Trim("" & txtGls_Materno.Text) & "','" & Trim("" & txtGls_Nombres.Text) & "','" & Trim("" & txt_RUC.Text) & "','" & Trim("" & TxtBrevete.Text) & "'"
        Cn.Execute StrCadSql
        
        Call fill
    End If
    
    FraRegChofer.Visible = False
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Activate()

    TxtBusq.SetFocus

End Sub

Private Sub Form_Deactivate()

    If EsNuevo = False Then
        TxtBusq.Text = ""
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
 
    Me.Caption = "Búsqueda"
    ConfGrid g, False, False, False, False
    EsNuevo = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    TxtBusq.Text = ""

End Sub

Private Sub g_OnDblClick()

    g_OnKeyDown 13, 1

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error Resume Next

    Select Case KeyCode
        Case 13:
            SRptBus(0) = g.Columns.ColumnByFieldName("cod").Value
           SRptBus(1) = g.Columns.ColumnByFieldName("des").Value
           g.Dataset.Close
           g.Dataset.Active = False
           Me.Hide
        Case 27
            Text1.SetFocus
    End Select

End Sub

Private Sub Text1_GotFocus()
    
    Unload Me

End Sub

Private Sub TxtBusq_Change()
    
    If EsNuevo = False Then
        CmdBusq_Click
    End If

End Sub

Private Sub TxtBusq_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyDown Then g.SetFocus
    If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub TxtBusq_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        CmdBusq_Click
        If g.Count > 1 Then g.SetFocus
    End If
    EsNuevo = False

End Sub

Private Sub fill()
Dim rsdatos                     As New ADODB.Recordset

sqlCond = Replace(sqlBus, "%X%", "%" & Trim(TxtBusq.Text) & "%") & SqlAdic & " order by 1"

If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open sqlCond, Cn, adOpenStatic, adLockOptimistic
    
Set g.DataSource = rsdatos
    
'    With G
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = sqlCond
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "cod"
'    End With
    LblReg.Caption = "(" + Format(g.Count, "0") + ")Registros"
    
End Sub

Public Sub Execute(strParAyuda As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    SRptBus(0) = ""
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
            cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
            cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
        Else
           cadRelacion = ""
           cadCondicion = ""
        End If
    Else
        cadRelacion = ""
        cadCondicion = ""
    End If
    
    SqlAdic = strParAdic
    strAyuda = ""
    strAyuda = strParAyuda
    sqlBus = setSql(strParAyuda)
    fill
    
    If strAyuda = "CHOFER" Or strAyuda = "VEHICULO" Then
        cmdopera(2).Visible = True
    End If
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        If TextBox2 Is Nothing Then
        Else
            TextBox2.Text = SRptBus(1)
        End If
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Private Function setSql(strParAyuda As String) As String
    
    setSql = ""
    Select Case UCase(strParAyuda)
         Case "CAMAL": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,proveedores WHERE personas.idPersona = proveedores.idProveedor AND proveedores.idEmpresa = '" & glsEmpresa & "' AND GlsPersona like '%X%' "
         Case "CLIENTE": setSql = "SELECT p.idPersona cod ,p.GlsPersona des FROM personas p Inner Join  clientes c  On  p.idPersona = c.idCliente AND c.idEmpresa = '" & glsEmpresa & "' " & cadRelacion & " WHERE  " & cadCondicion & "  (p.GlsPersona like '%X%' or p.idpersona like '%X%') "
         Case "PROVEEDOR": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,proveedores WHERE personas.idPersona = proveedores.idProveedor AND proveedores.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "VENDEDOR": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,vendedores WHERE personas.idPersona = vendedores.idVendedor AND vendedores.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona like '%X%' or personas.idpersona like '%X%')"
         Case "USUARIOS": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,usuarios WHERE personas.idPersona = usuarios.idUsuario AND usuarios.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "CHOFER": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,choferes WHERE personas.idPersona = choferes.idChofer AND choferes.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "EMPTRANS": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,EmpTrans WHERE personas.idPersona = EmpTrans.idEmpTrans AND EmpTrans.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "VEHICULO": setSql = "SELECT idVehiculo cod ,GlsVehiculo des FROM vehiculos WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsVehiculo like '%X%' or idVehiculo like '%X%')"
         Case "TIPOPERSONA": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '01' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "COBRADOR": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,cobradores WHERE personas.idPersona = Cobradores.idCobrador AND Cobradores.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "USUARIO": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,usuarios WHERE personas.idPersona = usuarios.idUsuario AND usuarios.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "AVAL": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,aval WHERE personas.idPersona = aval.idAval AND aval.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona  like '%X%' or personas.idpersona like '%X%')"
         Case "PAIS": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '02' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "DEPARTAMENTO": setSql = "SELECT idDpto cod, GlsUbigeo des FROM ubigeo WHERE idProv = '00' and idDist = '00' AND (GlsUbigeo like '%X%' or idDpto like '%X%')"
         Case "PROVINCIA": setSql = "SELECT idProv cod, GlsUbigeo des FROM ubigeo WHERE idProv <> '00' and idDist = '00' AND (GlsUbigeo like '%X%' or idProv like '%X%')"
         Case "DISTRITO": setSql = "SELECT idDistrito cod, GlsUbigeo des FROM ubigeo WHERE idDist <> '00' AND (GlsUbigeo like '%X%' or idDistrito like '%X%')"
         Case "PERSONA": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE (GlsPersona  like '%X%' or idpersona like '%X%')"
         Case "PERSONACLIENTE": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select c.idCliente from clientes c where c.idEmpresa = '" & glsEmpresa & "') AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONASUCURSAL": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select s.idSucursal from sucursales s where s.idEmpresa = '" & glsEmpresa & "') AND TipoPersona = '01002' AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONAPROVEEDOR": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select c.idProveedor from proveedores c  where c.idEmpresa = '" & glsEmpresa & "') AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONAVENDEDOR": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select v.idVendedor from vendedores v  where v.idEmpresa = '" & glsEmpresa & "') AND TipoPersona = '01001' AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONAUSUARIO": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select u.idUsuario from usuarios u  where u.idEmpresa = '" & glsEmpresa & "') AND TipoPersona = '01001' AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONACHOFER": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select c.idChofer from choferes c where c.idEmpresa = '" & glsEmpresa & "') AND TipoPersona = '01001' AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONAEMPTRANS": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select e.idEmpTrans from emptrans e where e.idEmpresa = '" & glsEmpresa & "') AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONACOBRADOR": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select e.idCobrador from cobradores e where e.idEmpresa = '" & glsEmpresa & "') AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "PERSONAAVAL": setSql = "SELECT idPersona cod ,GlsPersona des FROM personas WHERE idPersona not in (select e.idAval from aval e where e.idEmpresa = '" & glsEmpresa & "') AND (GlsPersona like '%X%' or idpersona like '%X%')"
         Case "LISTAPRECIOS": setSql = "SELECT l.idLista cod,l.GlsLista des FROM listaprecios l WHERE l.idEmpresa = '" & glsEmpresa & "' AND l.estLista = 1 AND (l.GlsLista like '%X%' or l.idLista like '%X%')"
         Case "NIVEL": setSql = "SELECT idNivel cod,GlsNivel des FROM niveles WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsNivel like '%X%' or idNivel like '%X%')"
         Case "MARCAVEHI": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '07' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TIPOPRODUCTO": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '06' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "MARCA": setSql = "SELECT idMarca cod,GlsMarca des FROM marcas WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsMarca like '%X%' or idMarca like '%X%')"
         Case "MONEDA": setSql = "SELECT idMoneda cod,GlsMoneda des FROM monedas WHERE  estMoneda = 'ACT' AND (GlsMoneda like '%X%' or idMoneda like '%X%')"
         Case "UM": setSql = "SELECT idUM cod,abreUM des FROM unidadmedida WHERE  (GlsUM like '%X%' or idUM like '%X%')"
         Case "UMGLOSA": setSql = "SELECT idUM cod,GlsUM des FROM unidadmedida WHERE  (GlsUM like '%X%' or idUM like '%X%')"
         Case "PRESENTACIONES": setSql = "SELECT presentaciones.idUM cod,abreUM des FROM presentaciones,unidadmedida WHERE presentaciones.idUM = unidadmedida.idUM AND presentaciones.idEmpresa = '" & glsEmpresa & "' AND (GlsUM like '%X%' or presentaciones.idUM like '%X%')"
         Case "ALMACEN": setSql = "SELECT idAlmacen cod,GlsAlmacen des FROM almacenes WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND (GlsAlmacen like '%X%' or idAlmacen like '%X%')"
         Case "ALMACENVTA": setSql = "SELECT idAlmacen cod,GlsAlmacen des FROM almacenes WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsAlmacen like '%X%' or idAlmacen like '%X%')"
         Case "MOTIVOTRASLADO": setSql = "SELECT idMotivoTraslado cod,GlsMotivoTraslado des FROM motivostraslados WHERE  (GlsMotivoTraslado like '%X%' or idMotivoTraslado like '%X%')"
         Case "TIPOFORMASPAGO": setSql = "SELECT idTipoFormaPago cod,GlsTipoFormaPago des FROM tipoformaspago WHERE  (GlsTipoFormaPago like '%X%' or idTipoFormaPago like '%X%')"
         Case "FORMASPAGO": setSql = "SELECT idFormaPago cod,GlsFormaPago des FROM formaspagos WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsFormaPago like '%X%' or idFormaPago like '%X%')"
         Case "DOCUMENTOS": setSql = "SELECT iddocumento cod,Glsdocumento des FROM documentos WHERE indVentas = '*' AND  (Glsdocumento like '%X%' or iddocumento like '%X%')"
         Case "DOCUMENTOSI": setSql = "SELECT iddocumento cod,Glsdocumento des FROM documentos WHERE (Glsdocumento like '%X%' or idDocumento like '%X%')"
         Case "MOTIVONCD": setSql = "SELECT idMotivoNCD cod,GlsMotivoNCD des FROM motivosncd WHERE (GlsMotivoNCD like '%X%' or idMotivoNCD like '%X%')"
         Case "PRODUCTOS": setSql = "SELECT idProducto cod,GlsProducto des FROM productos WHERE idEmpresa = '" & glsEmpresa & "' AND estProducto = 'A' AND (GlsProducto like '%X%' or idProducto like '%X%')"
         Case "CONCEPTOSALIDA": setSql = "SELECT idConcepto cod,glsConcepto des FROM conceptos WHERE tipoConcepto = 'S' AND (glsConcepto like '%X%' or idConcepto like '%X%')"
         Case "CONCEPTOINGRESO": setSql = "SELECT idConcepto cod,glsConcepto des FROM conceptos WHERE tipoConcepto = 'I' AND (glsConcepto like '%X%' or idConcepto like '%X%')"
         Case "PERFIL": setSql = "SELECT idPerfil cod,GlsPerfil des FROM perfil WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsPerfil like '%X%' or idPerfil like '%X%')"
         Case "EMPRESA": setSql = "SELECT idEmpresa cod,GlsEmpresa des FROM empresas WHERE (GlsEmpresa like '%X%' or idEmpresa like '%X%')"
         Case "SUCURSAL": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,sucursales WHERE personas.idPersona = sucursales.idSucursal AND sucursales.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona like '%X%' or personas.idpersona like '%X%') "
         Case "TIPONIVEL": setSql = "SELECT idTipoNivel cod,GlsTipoNivel des FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsTipoNivel like '%X%' or idTipoNivel like '%X%')"
         Case "CAJAS": setSql = "SELECT idCaja cod,GlsCaja des FROM cajas WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsCaja like '%X%' or idCaja like '%X%')"
         Case "CAJASUSUARIO": setSql = "SELECT c.idCaja cod,c.GlsCaja des FROM cajas c, cajasusuario u WHERE c.idCaja = u.idCaja AND c.idEmpresa = '" & glsEmpresa & "' AND u.idEmpresa = '" & glsEmpresa & "' AND u.idUsuario = '" & glsUser & "' AND (c.GlsCaja like '%X%' or c.idCaja like '%X%')"
         Case "CAJASUSUARIOFILTRO": setSql = "SELECT c.idCaja cod,c.GlsCaja des FROM cajas c, cajasusuario u WHERE c.idCaja = u.idCaja AND c.idEmpresa = '" & glsEmpresa & "' AND u.idEmpresa = '" & glsEmpresa & "' AND (c.GlsCaja like '%X%' or c.idCaja like '%X%')"
         Case "TIPOMOVCAJA": setSql = "SELECT c.idTipoMovCaja cod,c.GlsTipoMovCaja des FROM tiposmovcaja c WHERE LEFT(c.idTipoMovCaja,2) <> '99' AND (c.GlsTipoMovCaja like '%X%' or c.idTipoMovCaja like '%X%')"
         Case "CENTROCOSTO": setSql = "SELECT c.idCentroCosto cod,c.GlsCentroCosto des FROM centroscosto c WHERE c.idEmpresa = '" & glsEmpresa & "' AND (c.GlsCentroCosto like '%X%' or c.idCentroCosto like '%X%')"
         Case "DOCUMENOSEXP": setSql = "SELECT c.idDocumentoExp cod,d.GlsDocumento des FROM documentosexportar c,documentos d WHERE c.idDocumentoExp = d.idDocumento AND (d.GlsDocumento like '%X%' or c.idDocumentoExp like '%X%')"
         Case "PERMISOS": setSql = "SELECT idPermiso cod,GlsPermiso des FROM permisos WHERE (GlsPermiso like '%X%' or idPermiso like '%X%')"
         Case "CTACORRIENTE": setSql = "SELECT c.idCtaCorriente cod,c.GlsCtaCorriente des FROM ctascorrientes c WHERE c.idEmpresa = '" & glsEmpresa & "' AND (c.GlsCtaCorriente like '%X%' or c.idCtaCorriente like '%X%')"
         Case "BANCO": setSql = "SELECT c.idBanco cod,c.GlsBanco des FROM bancos c WHERE (c.GlsBanco like '%X%' or c.idBanco like '%X%')"
         Case "GRUPOSPRODUCTO": setSql = "SELECT idGrupo cod,GlsGrupo des FROM GruposProducto WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsGrupo like '%X%' or idGrupo like '%X%')"
         Case "TIPOTICKET": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '08' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "EMPRESA": setSql = "SELECT idEmpresa cod,GlsEmpresa des FROM empresas WHERE (GlsEmpresa like '%X%' or idEmpresa like '%X%')"
         Case "ZONAS": setSql = "SELECT idZona cod,GlsZona des FROM zonas WHERE (GlsZona like '%X%' or idZona like '%X%')"
         Case "UBICACIONES": setSql = "SELECT idUbicacion cod,GlsUbicacion des FROM almacenesubicacion WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsUbicacion like '%X%' or idUbicacion like '%X%')"
         Case "PERIODOLETRA": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '09' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TIPOINTERES": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '10' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TIPOCUOTALETRA": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '11' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TIPOCAPITALIZACION": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '12' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TIPOCOBRO": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '13' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TIPOEFECTIVO": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '14' AND (GlsDato like '%X%' or idDato like '%X%')"
         Case "TALLASPESOS": setSql = "SELECT c.idTallaPeso cod,c.GlsTallaPeso des FROM tallapeso c WHERE idEmpresa = '" & glsEmpresa & "' AND (GlsTallaPeso like '%X%' or c.idTallaPeso like '%X%')"
         Case "UNIDADPRODUC": setSql = "SELECT c.CodUnidProd cod,c.Descunidad des FROM unidadproduccion c WHERE idEmpresa = '" & glsEmpresa & "' AND (Descunidad like '%X%' or c.CodUnidProd like '%X%')"
         Case "CANALES": setSql = "SELECT idCanal cod, glsCanal des FROM Canal WHERE idEmpresa = '" & glsEmpresa & "' AND (glsCanal like '%X%' or idCanal like '%X%')"
         Case "TIPOSNIVEL": setSql = "SELECT idDato cod, glsDato des FROM datos WHERE idTipoDatos = '20' AND (glsDato like '%X%' or idDato like '%X%')"
         Case "ESTDOCUMENTOS": setSql = "SELECT distinct estDocVentas cod, CASE estDocVentas WHEN 'ANU' THEN 'ANULADO' WHEN 'CAN' THEN 'CANCELADO' WHEN 'GEN' THEN 'GENERADO' WHEN 'IMP' THEN 'IMPRESO' END  des FROM docventas WHERE (estDocVentas like '%X%')"
         Case "ZONA": setSql = "SELECT idZona cod, GlsZona des FROM Zonas WHERE (GlsZona like '%X%' or idZona like '%X%')"
         Case "CONTACTOSCLIENTES": setSql = "SELECT C.IDCONTACTO AS COD,P.GLSPERSONA AS DES FROM CONTACTOSCLIENTES AS C INNER JOIN PERSONAS P ON C.IDCONTACTO=P.IDPERSONA WHERE idEmpresa = '" & glsEmpresa & "' AND (P.GLSCONTACTO  like '%X%' or C.IDCONTACTO like '%X%')"
         Case "CLIENTES2": setSql = "SELECT p.idPersona cod ,p.GlsPersona des FROM personas p INNER JOIN clientes c ON c.idEmpresa = '" & glsEmpresa & "' AND p.idPersona = c.idCliente " & _
                                     "LEFT JOIN ubigeo u ON P.idDistrito = u.idDistrito " & _
                                     "LEFT JOIN ubigeo d ON left(u.idDistrito,2) = d.idDpto AND d.idProv = '00' AND d.idDist = '00' " & _
                                     "INNER JOIN personas v on c.idVendedorCampo = v.idPersona " & _
                                     "WHERE " & strVacio & "  c.idEmpresa = '" & glsEmpresa & "' AND (p.GlsPersona like '%X%' or p.idpersona like '%X%')"
                                                                     
          Case "FORMASPAGOXCLIENTE": setSql = "SELECT a.idformapago as cod,b.glsformapago as des FROM clientesformapagos A INNER JOIN formaspagos b on a.idformapago = b.idformapago and a.idempresa = b.idempresa AND (GlsFormaPago  like '%X%' or a.idFormaPago like '%X%')"
          Case "SUCURSALDESTINO": setSql = "SELECT idsucursal AS COD,glsabrev AS DES FROM Sucursales WHERE idEmpresa = '" & glsEmpresa & "' "
          Case "ALMACENDESTINO": setSql = "SELECT idalmacen AS COD,glsalmacen AS DES FROM Almacenes WHERE idEmpresa = '" & glsEmpresa & "' "
          Case "VENDEDORJEFE": setSql = "SELECT personas.idPersona cod ,GlsPersona des FROM personas,vendedores WHERE personas.idPersona = vendedores.idVendedor And indJefe = '1' And vendedores.idEmpresa = '" & glsEmpresa & "' AND (GlsPersona like '%X%' or personas.idpersona like '%X%')"
          Case "TIPODOCUMENTOIDENTIDAD": setSql = "SELECT IdTipoDocIdentidad cod ,GlsTipoDocIdentidad des FROM TiposDocIdentidad WHERE (IdTipoDocIdentidad like '%X%' or GlsTipoDocIdentidad like '%X%')"
          Case "CONCEPTOSCOSTEO": setSql = "SELECT idDato cod,GlsDato des FROM datos WHERE idtipoDatos = '26' AND (GlsDato like '%X%' or idDato like '%X%')"
          Case "CENTROCOSTOCLIENTE": setSql = "SELECT c.idCentroCosto cod,c.GlsCentroCosto des FROM centroscosto c WHERE (c.GlsCentroCosto like '%X%' or c.idCentroCosto like '%X%')"
          Case "LOTES": setSql = "SELECT IdLote cod, GlsLote des FROM Lotes WHERE idEmpresa = '" & glsEmpresa & "' And Estado = 'ACT' AND (IdLote like '%X%' or GlsLote like '%X%')"
          Case "DIRRECOJO": setSql = "SELECT c.iddirrecojo cod,c.GlsDirRecojo des FROM dirrecojos c WHERE C.idempresa = '" & glsEmpresa & "' and (c.GlsDirRecojo like '%X%' or c.iddirrecojo like '%X%')"
          Case "CONTACTOSPROVEEDORES": setSql = "SELECT C.IDCONTACTO AS COD,P.GLSPERSONA AS DES FROM contactosproveedores AS C INNER JOIN PERSONAS P ON C.IDCONTACTO=P.IDPERSONA WHERE idEmpresa = '" & glsEmpresa & "' AND (P.GLSCONTACTO like '%X%' or C.IDCONTACTO like '%X%') "
    End Select

End Function

Public Sub ExecuteReturnText(strParAyuda As String, ByRef strCod As String, ByRef strDes As String, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    SRptBus(0) = ""
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
    If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
            cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
            cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
        Else
           cadRelacion = ""
           cadCondicion = ""
        End If
    Else
        cadRelacion = ""
        cadCondicion = ""
    End If
    
    SqlAdic = strParAdic
    strAyuda = ""
    strAyuda = strParAyuda
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Public Sub ExecuteKeyascii(ByVal KeyAscii As Integer, strParAyuda As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
Dim StrMsgError As String
    
    MousePointer = 0
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
            cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
            cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
        Else
           cadRelacion = ""
           cadCondicion = ""
        End If
    Else
        cadRelacion = ""
        cadCondicion = ""
    End If
    
    SqlAdic = strParAdic
    strAyuda = ""
    strAyuda = strParAyuda
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Public Sub ExecuteKeyasciiReturnText(ByVal KeyAscii As Integer, strParAyuda As String, ByRef strCod As String, ByRef strDes As String, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    SRptBus(0) = ""
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
            cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
            cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
        Else
           cadRelacion = ""
           cadCondicion = ""
        End If
    Else
        cadRelacion = ""
        cadCondicion = ""
    End If
    
    SqlAdic = strParAdic
    strAyuda = ""
    strAyuda = strParAyuda
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Private Sub txtCod_MarcaVehi_Change()
    txtGls_MarcaVehi.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_MarcaVehi.Text, False)

End Sub

