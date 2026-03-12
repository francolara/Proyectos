VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmResumenVentas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Ventas"
   ClientHeight    =   8505
   ClientLeft      =   1620
   ClientTop       =   1545
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   14325
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7755
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   14235
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   14010
         Begin VB.TextBox txtAño 
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
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   16
            Top             =   450
            Width           =   675
         End
         Begin VB.Frame fraReportes 
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   765
            Index           =   13
            Left            =   2430
            TabIndex        =   5
            Top             =   1035
            Width           =   11235
            Begin VB.CommandButton cmbAyudaCliente 
               Height          =   315
               Left            =   10725
               Picture         =   "FrmResumenVentas.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   285
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Cliente 
               Height          =   315
               Left            =   765
               TabIndex        =   7
               Tag             =   "TidPerCliente"
               Top             =   270
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
               Container       =   "FrmResumenVentas.frx":038A
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Cliente 
               Height          =   315
               Left            =   1710
               TabIndex        =   8
               Tag             =   "TGlsCliente"
               Top             =   270
               Width           =   9000
               _ExtentX        =   15875
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
               Locked          =   -1  'True
               Container       =   "FrmResumenVentas.frx":03A6
               Estilo          =   1
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtdir 
               Height          =   285
               Left            =   10620
               TabIndex        =   9
               Tag             =   "TidPerCliente"
               Top             =   90
               Visible         =   0   'False
               Width           =   195
               _ExtentX        =   344
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
               Container       =   "FrmResumenVentas.frx":03C2
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtruc 
               Height          =   285
               Left            =   10845
               TabIndex        =   10
               Tag             =   "TidPerCliente"
               Top             =   90
               Visible         =   0   'False
               Width           =   150
               _ExtentX        =   265
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
               Container       =   "FrmResumenVentas.frx":03DE
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCodtienda 
               Height          =   285
               Left            =   0
               TabIndex        =   21
               Tag             =   "TidPerCliente"
               Top             =   0
               Visible         =   0   'False
               Width           =   150
               _ExtentX        =   265
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
               Container       =   "FrmResumenVentas.frx":03FA
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin VB.Label lbl_Cliente 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
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
               Height          =   210
               Left            =   180
               TabIndex        =   11
               Top             =   315
               Width           =   480
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            Caption         =   "Moneda"
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
            Height          =   735
            Left            =   8010
            TabIndex        =   4
            Top             =   180
            Width           =   5640
            Begin VB.OptionButton OptOri 
               Caption         =   "Original"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3420
               TabIndex        =   20
               Top             =   360
               Width           =   1230
            End
            Begin VB.OptionButton OptConversion 
               Caption         =   "Conversion"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1215
               TabIndex        =   19
               Top             =   315
               Width           =   1320
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            Caption         =   "Rango de Meses"
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
            Height          =   735
            Left            =   2430
            TabIndex        =   3
            Top             =   180
            Width           =   5280
            Begin VB.ComboBox cbx_MesDes 
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
               ItemData        =   "FrmResumenVentas.frx":0416
               Left            =   765
               List            =   "FrmResumenVentas.frx":0418
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   270
               Width           =   1710
            End
            Begin VB.ComboBox cbx_MesHas 
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
               ItemData        =   "FrmResumenVentas.frx":041A
               Left            =   3240
               List            =   "FrmResumenVentas.frx":041C
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   270
               Width           =   1710
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   2655
               TabIndex        =   15
               Top             =   315
               Width           =   420
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   180
               TabIndex        =   14
               Top             =   315
               Width           =   465
            End
         End
         Begin MSComCtl2.UpDown udAño 
            Height          =   315
            Left            =   1845
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   450
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   765
            TabIndex        =   18
            Top             =   495
            Width           =   300
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   5295
         Left            =   90
         OleObjectBlob   =   "FrmResumenVentas.frx":041E
         TabIndex        =   1
         Top             =   2340
         Width           =   14040
      End
   End
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
            Picture         =   "FrmResumenVentas.frx":42E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":4683
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":4AD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":4E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":5209
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":55A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":593D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":5CD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":6071
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":640B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":67A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":7467
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":7801
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":7C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":7FED
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResumenVentas.frx":89FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   1164
      ButtonWidth     =   4260
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualizar"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Det por Cuenta"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Det por Documento"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmResumenVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAyudaCliente_Click()
    
    mostrarAyudaClientes txtCod_Cliente, txtGls_Cliente, txtruc, txtdir, txtCodtienda

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError             As String
Dim ameses(1 To 12)         As String * 35
Dim X                       As Long

    Toolbar1.Buttons(3).Visible = False

    ameses(1) = " Enero                         "
    ameses(2) = " Febrero                       "
    ameses(3) = " Marzo                         "
    ameses(4) = " Abril                         "
    ameses(5) = " Mayo                          "
    ameses(6) = " Junio                         "
    ameses(7) = " Julio                         "
    ameses(8) = " Agosto                        "
    ameses(9) = " Setiembre                     "
    ameses(10) = " Octubre                       "
    ameses(11) = " Noviembre                     "
    ameses(12) = " Diciembre                     "
    
    For X = 1 To 12
       cbx_MesDes.AddItem ameses(X)
       cbx_MesHas.AddItem ameses(X)
    Next X
    cbx_MesDes.ListIndex = 0
    cbx_MesHas.ListIndex = Val(Month(Format(Now, "dd/MM/yyyy"))) - 1
    
    txtAño.Text = Format$(Now, "yyyy")
    OptConversion.Value = True
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String
Dim CGlsReporte As String
Dim GlsForm     As String
Dim Tipo        As String
Dim strMesIni   As Integer
Dim strMesFin   As Integer

    If OptConversion.Value = True Then
        Tipo = "1"
    Else
        Tipo = "2"
    End If
    
    strMesIni = Val(cbx_MesDes.ListIndex + 1)
    strMesFin = Val(cbx_MesHas.ListIndex + 1)

    Select Case Button.Index
        Case 1 'Actualizar
            actualizar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'EXCEL
            G.m.ExportToXLS App.Path & "\Temporales\ResumenVentas.xls"
            ShellEx App.Path & "\Temporales\ResumenVentas.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 3 'imp1
            CGlsReporte = "rptResumenVentasDetCuenta.rpt"
            GlsForm = "Reporte de Resumen de Ventas - Detalle por Cuenta Contable"
            mostrarReporte CGlsReporte, "parEmpresa|parTipo|parAno|parMesDesde|parMesHasta|parCliente", glsEmpresa & "|" & Tipo & "|" & txtAño.Text & "|" & strMesIni & "|" & strMesFin & "|" & Trim("" & txtCod_Cliente.Text), GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 4 'imp1
             CGlsReporte = "rptResumenVentasDetDocumento.rpt"
             GlsForm = "Reporte de Resumen de ventas - Detalle por Documento"
             
             mostrarReporte CGlsReporte, "varEmpresa|varTipo|varAno|varMesDesde|varMesHasta|varCliente", glsEmpresa & "|" & Tipo & "|" & txtAño.Text & "|" & strMesIni & "|" & strMesFin & "|" & Trim("" & txtCod_Cliente.Text), GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 5 'Salir
            Unload Me
    End Select

    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub actualizar(StrMsgError As String)
Dim Tipo        As String
Dim strMesIni   As Integer
Dim strMesFin   As Integer
Dim rsdatos     As New ADODB.Recordset
On Error GoTo Err

    If OptConversion.Value = True Then
        Tipo = "1"
    Else
        Tipo = "2"
    End If
    
    strMesIni = Val(cbx_MesDes.ListIndex + 1)
    strMesFin = Val(cbx_MesHas.ListIndex + 1)
       
    csql = "spu_ResumenVentas '" & glsEmpresa & "','" & Tipo & "'," & txtAño & "," & strMesIni & "," & strMesFin & ",'" & Trim("" & txtCod_Cliente.Text) & "'"
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set G.DataSource = rsdatos

'    With G
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = "CALL spu_ResumenVentas ('" & glsEmpresa & "','" & Tipo & "'," & txtAño & "," & strMesIni & "," & strMesFin & ",'" & Trim("" & txtCod_Cliente.Text) & "')"
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idPerCliente"
'    End With
'    Me.Refresh
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Cliente_Change()

    If txtCod_Cliente.Text <> "" Then
        If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
            If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
                txtGls_Cliente.Text = traerCampo("personas p Inner Join  clientes c On  p.idPersona = c.idCliente Inner Join personas v On c.idVendedorCampo = v.idPersona", "p.GlsPersona", "p.idPersona", txtCod_Cliente.Text, False, "c.idVendedorCampo ='" & glsUser & "' ")
            Else
                txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
            End If
        Else
            txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
        End If
    Else
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If

End Sub

Private Sub txtAño_Change()

    If Not (Val(txtAño.Text) > 0 And Val(txtAño.Text) < 10000) Then
        txtAño.Text = txtAño.Text
    End If
    
End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And (Not KeyAscii = 8) Then
       KeyAscii = 0
    End If

End Sub

Private Sub txtAño_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Then
        If Val(txtAño.Text) + 1 < 10000 Then
            txtAño.Text = Format(Val(txtAño.Text) + 1, "0000")
        End If
    End If
    If KeyCode = 40 Then
        If Val(txtAño.Text) - 1 > 0 Then
            txtAño.Text = Format(Val(txtAño.Text) - 1, "0000")
        End If
    End If
    
End Sub

Private Sub udAño_DownClick()

    If Val(txtAño.Text) - 1 > 0 Then
        txtAño.Text = Format(Val(txtAño.Text) - 1, "0000")
    End If

End Sub

Private Sub udAño_UpClick()

    If Val(txtAño.Text) - 1 < 10000 Then
        txtAño.Text = Format(Val(txtAño.Text) + 1, "0000")
    End If
    
End Sub
