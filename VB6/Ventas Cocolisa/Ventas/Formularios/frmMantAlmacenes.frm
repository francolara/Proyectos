VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantAlmacenes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Almacenes"
   ClientHeight    =   7620
   ClientLeft      =   3000
   ClientTop       =   1965
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10590
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   1620
      Top             =   7200
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
            Picture         =   "frmMantAlmacenes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenes.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
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
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   6855
      Left            =   45
      TabIndex        =   11
      Top             =   630
      Width           =   10455
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         TabIndex        =   12
         Top             =   150
         Width           =   10200
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   255
            Width           =   9075
            _ExtentX        =   16007
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
            Container       =   "frmMantAlmacenes.frx":3518
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
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5700
         Left            =   120
         OleObjectBlob   =   "frmMantAlmacenes.frx":3534
         TabIndex        =   1
         Top             =   1005
         Width           =   10200
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6825
      Left            =   45
      TabIndex        =   9
      Top             =   645
      Width           =   10440
      Begin TabDlg.SSTab SSTab1 
         Height          =   6315
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   11139
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "frmMantAlmacenes.frx":55C4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Ubicaciones"
         TabPicture(1)   =   "frmMantAlmacenes.frx":55E0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdgrabaubicacion"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame4"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame6 
            Height          =   3390
            Left            =   360
            TabIndex        =   23
            Top             =   810
            Width           =   9330
            Begin VB.CommandButton cmbAyudaSucursal 
               Height          =   315
               Left            =   8550
               Picture         =   "frmMantAlmacenes.frx":55FC
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   1260
               Width           =   390
            End
            Begin CATControls.CATTextBox txtGls_Almacen 
               Height          =   315
               Left            =   1290
               TabIndex        =   4
               Tag             =   "TglsAlmacen"
               Top             =   1695
               Width           =   7680
               _ExtentX        =   13547
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
               Container       =   "frmMantAlmacenes.frx":5986
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Sucursal 
               Height          =   315
               Left            =   1290
               TabIndex        =   3
               Tag             =   "TidSucursal"
               Top             =   1260
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
               Container       =   "frmMantAlmacenes.frx":59A2
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Sucursal 
               Height          =   315
               Left            =   2265
               TabIndex        =   25
               Top             =   1260
               Width           =   6270
               _ExtentX        =   11060
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
               Container       =   "frmMantAlmacenes.frx":59BE
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Almacen 
               Height          =   315
               Left            =   8025
               TabIndex        =   2
               Tag             =   "TidAlmacen"
               Top             =   450
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
               Container       =   "frmMantAlmacenes.frx":59DA
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               Caption         =   "Sucursal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   240
               Left            =   270
               TabIndex        =   28
               Top             =   1305
               Width           =   765
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
               TabIndex        =   27
               Top             =   1755
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
               Left            =   7380
               TabIndex        =   26
               Top             =   480
               Width           =   495
            End
         End
         Begin VB.CommandButton cmdgrabaubicacion 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74730
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1845
            Width           =   1140
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1230
            Left            =   -74730
            TabIndex        =   20
            Top             =   495
            Width           =   9555
            Begin CATControls.CATTextBox TxtGls_Ubicacion 
               Height          =   315
               Left            =   1350
               TabIndex        =   6
               Top             =   675
               Width           =   7905
               _ExtentX        =   13944
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
               MaxLength       =   100
               Container       =   "frmMantAlmacenes.frx":59F6
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Ubicacion 
               Height          =   315
               Left            =   1350
               TabIndex        =   5
               Top             =   270
               Width           =   2730
               _ExtentX        =   4815
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
               Container       =   "frmMantAlmacenes.frx":5A12
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label2 
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
               TabIndex        =   22
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label5 
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
               Left            =   270
               TabIndex        =   21
               Top             =   315
               Width           =   495
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   3795
            Left            =   -74730
            TabIndex        =   19
            Top             =   2250
            Width           =   9555
            Begin DXDBGRIDLibCtl.dxDBGrid gUbicaciones 
               Height          =   3405
               Left            =   135
               OleObjectBlob   =   "frmMantAlmacenes.frx":5A2E
               TabIndex        =   8
               Top             =   225
               Width           =   9270
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   4230
            Left            =   -74460
            TabIndex        =   17
            Top             =   855
            Width           =   9690
            Begin DXDBGRIDLibCtl.dxDBGrid gPresentaciones 
               Height          =   3765
               Left            =   165
               OleObjectBlob   =   "frmMantAlmacenes.frx":75F1
               TabIndex        =   18
               Top             =   270
               Width           =   9420
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   6885
            Left            =   -74550
            TabIndex        =   15
            Top             =   810
            Width           =   9780
            Begin DXDBGRIDLibCtl.dxDBGrid gAlmacenes 
               Height          =   6420
               Left            =   135
               OleObjectBlob   =   "frmMantAlmacenes.frx":96B7
               TabIndex        =   16
               Top             =   270
               Width           =   9510
            End
         End
      End
   End
End
Attribute VB_Name = "frmMantAlmacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sw_ModUbicacion As Boolean

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmdgrabaubicacion_Click()
On Error GoTo Err
Dim StrMsgError     As String
Dim Cadmysql        As String
    
    If Len(Trim("" & txtCod_Almacen.Text)) = 0 Then
        StrMsgError = "El almacen debe existir para agregar las ubicaciones"
        GoTo Err
    End If
    
    If Len(Trim("" & txtCod_Ubicacion.Text)) = 0 Then
        StrMsgError = "Debe Colocar un Codigo"
        GoTo Err
    End If
    
    If sw_ModUbicacion = False Then
        If Len(Trim("" & traerCampo("almacenesubicacion", "idUbicacion", "idUbicacion", Trim(txtCod_Ubicacion.Text), True, " idalmacen = '" & txtCod_Almacen.Text & "' "))) > 0 Then
            StrMsgError = "La ubicación ya existe en el almacen"
            GoTo Err
        End If
        
        If Len(Trim("" & traerCampo("almacenesubicacion", "idUbicacion", "idUbicacion", Trim(txtCod_Ubicacion.Text), True))) > 0 Then
            StrMsgError = "La ubicación ya existe en Otro almacen"
            GoTo Err
        End If
        
        Cadmysql = "Insert Into almacenesubicacion(idUbicacion, GlsUbicacion, idAlmacen, idEmpresa) " & _
                    "Values('" & Trim(txtCod_Ubicacion.Text) & "','" & Trim(TxtGls_Ubicacion.Text) & "','" & txtCod_Almacen.Text & "','" & glsEmpresa & "')"
        Cn.Execute (Cadmysql)
        
    Else
        Cadmysql = "Update almacenesubicacion set GlsUbicacion = '" & Trim("" & TxtGls_Ubicacion.Text) & "'" & _
                   "where idalmacen = '" & Trim("" & txtCod_Almacen.Text) & "' and idubicacion = '" & Trim("" & txtCod_Ubicacion.Text) & "' and idempresa = '" & glsEmpresa & "' "
        Cn.Execute (Cadmysql)
        
        txtCod_Ubicacion.Enabled = True
        txtCod_Ubicacion.BackColor = &HFFFFFF
        
        sw_ModUbicacion = False
    End If
    
    txtCod_Ubicacion.Text = ""
    TxtGls_Ubicacion.Text = ""
    
    listaUbicaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    ConfGrid gUbicaciones, True, False, False, False
    
    listaAlmacen StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo
    sw_ModUbicacion = False
    
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
    
    validaHomonimia "almacenes", "GlsAlmacen", "idAlmacen", txtGls_Almacen.Text, txtCod_Almacen.Text, True, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Almacen.Text = "" Then
        txtCod_Almacen.Text = GeneraCorrelativoAnoMes("almacenes", "idAlmacen")
        EjecutaSQLForm Me, 0, True, "almacenes", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabó"
    
    Else
        EjecutaSQLForm Me, 1, True, "almacenes", StrMsgError, "idAlmacen"
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Modificó"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    
    listaAlmacen StrMsgError
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

    mostrarAlmacen gLista.Columns.ColumnByName("idAlmacen").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnReloadGroupList()
    
    gLista.m.FullExpand

End Sub

Private Sub gUbicaciones_OnDblClick()
On Error GoTo Err
Dim StrMsgError     As String
Dim Cadmysql        As String
Dim RsUbicacion     As New ADODB.Recordset

    If RsUbicacion.State = 1 Then RsUbicacion.Close
    Set RsUbicacion = Nothing

    Cadmysql = "Select idUbicacion, GlsUbicacion, idAlmacen, idEmpresa from almacenesubicacion " & _
               "where idalmacen = '" & Trim("" & txtCod_Almacen.Text) & "' and idubicacion = '" & Trim("" & gUbicaciones.Columns.ColumnByFieldName("idubicacion").Value) & "' and idempresa = '" & glsEmpresa & "' "

    RsUbicacion.Open Cadmysql, Cn, adOpenStatic, adLockOptimistic
    
    If Not RsUbicacion.EOF Then
        txtCod_Ubicacion.Text = Trim("" & RsUbicacion.Fields("idUbicacion"))
        TxtGls_Ubicacion.Text = Trim("" & RsUbicacion.Fields("GlsUbicacion"))
        
        sw_ModUbicacion = True
        txtCod_Ubicacion.Enabled = False
        txtCod_Ubicacion.BackColor = &HC0FFFF
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gUbicaciones_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim StrMsgError     As String
Dim i               As Integer
Dim Cadmysql        As String

    If KeyCode = 46 Then
        If gUbicaciones.Count > 0 Then
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If Len(Trim("" & traerCampo("productosalmacen", "idUbicacion", "idUbicacion", Trim("" & gUbicaciones.Columns.ColumnByFieldName("idUbicacion").Value), True, " idalmacen = '" & Trim("" & txtCod_Almacen.Text) & "'"))) > 0 Then
                    StrMsgError = "La Ubicacion ya esta siendo usada ,Imposible eliminar"
                    GoTo Err
                End If
                
                Cadmysql = "Delete from almacenesubicacion where idempresa = '" & glsEmpresa & "' " & _
                           "and idUbicacion = '" & Trim("" & gUbicaciones.Columns.ColumnByFieldName("idUbicacion").Value) & "' and idalmacen = '" & Trim("" & txtCod_Almacen.Text) & "' "
                Cn.Execute (Cadmysql)
                
                txtCod_Ubicacion.Text = ""
                TxtGls_Ubicacion.Text = ""
                
                listaUbicaciones StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
        End If
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
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

    listaAlmacen StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaAlmacen(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsAlmacen LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT a.idSucursal,s.GlsPersona AS GlsSucursal, a.idAlmacen, a.GlsAlmacen " & _
           "FROM almacenes a,personas s WHERE a.idSucursal = s.idPersona AND a.idEmpresa = '" & glsEmpresa & "'"
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY a.idAlmacen"
    

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
'        .KeyField = "idAlmacen"
'    End With
    gLista.Columns.ColumnByName("GlsSucursal").GroupIndex = 0
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarAlmacen(strCodAlm As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT a.idAlmacen,a.GlsAlmacen,a.idSucursal " & _
           "FROM almacenes a " & _
           "WHERE a.idAlmacen = '" & strCodAlm & "' AND a.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    listaUbicaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Sucursal_Change()
    
    txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminarán todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCod_Almacen.Text)
    
    csql = "SELECT idAlmacen FROM docventas WHERE idAlmacen = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Ventas)."
        GoTo Err
    End If
    
    csql = "SELECT idAlmacen FROM valescab WHERE idAlmacen = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Vales)."
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando ventas almacen
    csql = "DELETE FROM AlmacenesVtas WHERE idAlmacen = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando productosalmacen
    csql = "DELETE FROM productosalmacen WHERE idAlmacen = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando el registro
    csql = "DELETE FROM almacenes WHERE idAlmacen = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    If rsValida.State = 1 Then rsValida.Close:  Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaUbicaciones(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond     As String
Dim Cadmysql    As String
Dim rsdatos                     As New ADODB.Recordset

    Cadmysql = "SELECT (@i:=@i +1) Item,idUbicacion, GlsUbicacion, idAlmacen, idEmpresa " & _
               "FROM almacenesubicacion ,(SELECT @i:= 0) foo WHERE idalmacen = '" & txtCod_Almacen.Text & "' and idEmpresa = '" & glsEmpresa & "' " & _
               "Order by idUbicacion "

If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
Set gUbicaciones.DataSource = rsdatos

'    With gUbicaciones
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = Cadmysql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idAlmacen"
'    End With
    
    
    
    gLista.Columns.ColumnByName("GlsSucursal").GroupIndex = 0
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

