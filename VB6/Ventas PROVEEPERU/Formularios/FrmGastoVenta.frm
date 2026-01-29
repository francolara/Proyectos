VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmGastoVenta 
   Caption         =   "Gasto de Ventas"
   ClientHeight    =   7035
   ClientLeft      =   3090
   ClientTop       =   2175
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9765
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
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
            Object.ToolTipText     =   "Eliminar"
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
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   7335
         Top             =   90
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
               Picture         =   "FrmGastoVenta.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":039A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":07EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":0B86
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":0F20
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":12BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":1654
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":19EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":1D88
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":2122
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":24BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGastoVenta.frx":317E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00000000&
      Height          =   6405
      Left            =   0
      TabIndex        =   1
      Top             =   615
      Width           =   9705
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   135
         TabIndex        =   2
         Top             =   195
         Width           =   9495
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1020
            TabIndex        =   3
            Top             =   210
            Width           =   8430
            _ExtentX        =   14870
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
            MaxLength       =   255
            Container       =   "FrmGastoVenta.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   165
            TabIndex        =   4
            Top             =   255
            Width           =   765
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5280
         Left            =   120
         OleObjectBlob   =   "FrmGastoVenta.frx":3534
         TabIndex        =   5
         Top             =   960
         Width           =   9495
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6405
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   9690
      Begin VB.Frame fraDetalle 
         Appearance      =   0  'Flat
         Caption         =   " Detalle "
         ForeColor       =   &H00000000&
         Height          =   3165
         Left            =   120
         TabIndex        =   7
         Top             =   3090
         Width           =   9375
         Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
            Height          =   2895
            Left            =   45
            OleObjectBlob   =   "FrmGastoVenta.frx":55DE
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   195
            Width           =   9225
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2880
         Left            =   105
         TabIndex        =   8
         Top             =   210
         Width           =   9390
         Begin VB.CommandButton cmbAyudaVendedor 
            Height          =   345
            Left            =   8910
            Picture         =   "FrmGastoVenta.frx":84DF
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1410
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_GastoVenta 
            Height          =   315
            Left            =   8355
            TabIndex        =   10
            Tag             =   "TidCodGastoVenta"
            Top             =   135
            Width           =   915
            _ExtentX        =   1614
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
            MaxLength       =   8
            Container       =   "FrmGastoVenta.frx":8869
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Descripcion 
            Height          =   600
            Left            =   1500
            TabIndex        =   11
            Tag             =   "TglsDescGastoVenta"
            Top             =   2100
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   1058
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
            MaxLength       =   255
            Container       =   "FrmGastoVenta.frx":8885
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   1485
            TabIndex        =   12
            Tag             =   "FFecRegistro"
            Top             =   450
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42598401
            CurrentDate     =   38667
         End
         Begin CATControls.CATTextBox txt_Monto 
            Height          =   285
            Left            =   1485
            TabIndex        =   13
            Tag             =   "NTotalGastoVenta"
            Top             =   930
            Width           =   1305
            _ExtentX        =   2302
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmGastoVenta.frx":88A1
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Vendedor 
            Height          =   315
            Left            =   1485
            TabIndex        =   14
            Tag             =   "TidVendedor"
            Top             =   1440
            Width           =   1305
            _ExtentX        =   2302
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
            Container       =   "FrmGastoVenta.frx":88BD
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Vendedor 
            Height          =   315
            Left            =   2865
            TabIndex        =   15
            Tag             =   "TglsVendedor"
            Top             =   1440
            Width           =   6000
            _ExtentX        =   10583
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
            Container       =   "FrmGastoVenta.frx":88D9
            Vacio           =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Descripcion:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   2310
            Width           =   885
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7620
            TabIndex        =   19
            Top             =   165
            Width           =   660
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   165
            TabIndex        =   18
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Monto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1500
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "FrmGastoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAyudaVendedor_Click()
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err
   
    ConfGrid GLista, False, False, False, False
    ConfGrid GDetalle, True, True, False, True
    
    listaGastoVenta StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    FraListado.Visible = True
    FraGeneral.Visible = False
    habilitaBotones 7
    
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnDblClick()
Dim StrMsgError As String
On Error GoTo Err

    mostrarGastoVenta GLista.Columns.ColumnByName("idCodGastoVenta").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    FraListado.Visible = False
    FraGeneral.Visible = True
    FraGeneral.Enabled = False
    habilitaBotones 2
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
    Select Case Button.Index
        Case 1 'Nuevo
            limpiaForm Me
            nuevo StrMsgError
            FraListado.Visible = False
            FraGeneral.Visible = True
            FraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            FraGeneral.Enabled = False
        Case 3 'Modificar
            FraGeneral.Enabled = True
        Case 4  'Cancelar
            FraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            FraGeneral.Enabled = True
        Case 6 'Imprimir
            GLista.m.ExportToXLS App.Path & "\Temporales\ListadoGastos.xls"
            ShellEx App.Path & "\Temporales\ListadoGastos.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 7 'Lista
            FraListado.Visible = True
            FraGeneral.Visible = False
        Case 8 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    Exit Sub
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
Private Sub habilitaBotones(indexBoton As Integer)

Select Case indexBoton
    Case 1, 5 'Nuevo, Eliminar
        Toolbar1.Buttons(1).Visible = False 'Nuevo
        Toolbar1.Buttons(2).Visible = True 'Grabar
        Toolbar1.Buttons(3).Visible = False 'Modificar
        Toolbar1.Buttons(4).Visible = False 'Cancelar
        Toolbar1.Buttons(5).Visible = False 'Eliminar
        Toolbar1.Buttons(6).Visible = False 'Imprimir
        Toolbar1.Buttons(7).Visible = True 'Lista
    Case 2, 4 'Grabar, Cancelar
        Toolbar1.Buttons(1).Visible = True 'Nuevo
        Toolbar1.Buttons(2).Visible = False 'Grabar
        Toolbar1.Buttons(3).Visible = True 'Modificar
        Toolbar1.Buttons(4).Visible = False 'Cancelar
        Toolbar1.Buttons(5).Visible = True 'Eliminar
        Toolbar1.Buttons(6).Visible = False 'Imprimir
        Toolbar1.Buttons(7).Visible = True 'Lista
    Case 3 'Modificar
        Toolbar1.Buttons(1).Visible = False 'Nuevo
        Toolbar1.Buttons(2).Visible = True 'Grabar
        Toolbar1.Buttons(3).Visible = False 'Modificar
        Toolbar1.Buttons(4).Visible = True 'Cancelar
        Toolbar1.Buttons(5).Visible = False 'Eliminar
        Toolbar1.Buttons(6).Visible = False 'Imprimir
        Toolbar1.Buttons(7).Visible = False 'Lista
    Case 7 'Lista
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
        Toolbar1.Buttons(5).Visible = False
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
End Select

End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
    listaGastoVenta StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
Private Sub listaGastoVenta(ByRef StrMsgError As String)
Dim strCond As String
On Error GoTo Err
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " And glsVendedor LIKE '%" & strCond & "%'"
    End If
    
    csql = "Select idCodGastoVenta, glsVendedor,glsDescGastoVenta,TotalGastoVenta,DATE_FORMAT(FecRegistro,'%d/%m/%Y') as FecRegistro From GastosVenta Where idEmpresa = '" & glsEmpresa & "' "
           
    If strCond <> "" Then csql = csql + strCond

    csql = csql + " ORDER BY idCodGastoVenta"
    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idCodGastoVenta"
    End With

    Me.Refresh
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
Dim StrCodigo As String
Dim strMsg      As String
On Error GoTo Err

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
   
    If txtCod_GastoVenta.Text = "" Then 'graba
        txtCod_GastoVenta.Text = GeneraCorrelativoAnoMes("GastosVenta", "idCodGastoVenta")
        
        EjecutaSQLForm Me, 0, True, "GastosVenta", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        EjecutaDetalleGasto StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Grabo"
    Else 'modifica
    
        EjecutaSQLForm Me, 1, True, "GastosVenta", StrMsgError, "idCodGastoVenta"
        If StrMsgError <> "" Then GoTo Err
        
        EjecutaDetalleGasto StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    FraGeneral.Enabled = False
    
    listaGastoVenta StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    'Resume
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
Dim indTrans As Boolean
Dim StrCodigo As String
Dim rsValida As New ADODB.Recordset

On Error GoTo Err

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    StrCodigo = Trim(txtCod_GastoVenta.Text)
 
    Cn.BeginTrans
    indTrans = True
    
    '*********************************************ELIMINANDO****************************************************
    'Eliminando Registro
    csql = "DELETE FROM GastosVenta WHERE idCodGastoVenta = '" & StrCodigo & "' AND idEmpresa = '" & glsEmpresa & " ' "
    Cn.Execute csql
    
    csql = "DELETE FROM GastosVentadet WHERE idCodGastoVenta = '" & StrCodigo & "' AND idEmpresa = '" & glsEmpresa & " ' "
    Cn.Execute csql
    
    Cn.CommitTrans
    
    'Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    
    listaGastoVenta StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If rsValida.State = 1 Then rsValida.Close
    Set rsValida = Nothing
    Exit Sub
Err:
    If rsValida.State = 1 Then rsValida.Close
    Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarGastoVenta(strCodGastoVenta As String, ByRef StrMsgError As String)
Dim rst  As New ADODB.Recordset
Dim rsd  As New ADODB.Recordset
On Error GoTo Err

    csql = "Select idCodGastoVenta, glsDescGastoVenta, idVendedor,glsVendedor, TotalGastoVenta,FecRegistro From GastosVenta " & _
           "Where idCodGastoVenta = '" & strCodGastoVenta & "' AND idEmpresa = '" & glsEmpresa & "' "
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    csql = "Select item,idCliente, ruc, NroDocumento, ValImporte, GlsCliente From GastosVentaDet Where idCodGastoVenta = '" & strCodGastoVenta & "' AND idEmpresa = '" & glsEmpresa & "' "

    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    rsd.Fields.Append "Item", adInteger, , adFldRowID
    rsd.Fields.Append "idCliente", adChar, 8, adFldIsNullable
    rsd.Fields.Append "ruc", adVarChar, 12, adFldIsNullable
    rsd.Fields.Append "NroDocumento", adChar, 100, adFldIsNullable
    rsd.Fields.Append "GlsCliente", adChar, 200, adFldIsNullable
    rsd.Fields.Append "ValImporte", adDouble, 14, adFldIsNullable
    rsd.Open
    
    If rst.RecordCount = 0 Then
        rsd.AddNew
        rsd.Fields("Item") = 1
        rsd.Fields("idCliente") = ""
        rsd.Fields("ruc") = ""
        rsd.Fields("NroDocumento") = ""
        rsd.Fields("GlsCliente") = ""
        rsd.Fields("ValImporte") = 0
      
    Else
        Do While Not rst.EOF
            rsd.AddNew
            rsd.Fields("Item") = "" & rst.Fields("Item")
            rsd.Fields("idCliente") = "" & rst.Fields("idCliente")
            rsd.Fields("ruc") = "" & rst.Fields("ruc")
            rsd.Fields("NroDocumento") = "" & rst.Fields("NroDocumento")
            rsd.Fields("GlsCliente") = "" & rst.Fields("GlsCliente")
            rsd.Fields("ValImporte") = "" & rst.Fields("ValImporte")
            rst.MoveNext
        Loop
    End If
    
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL GDetalle, rsd, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Refresh
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    'Resume
End Sub


Private Sub txtCod_Vendedor_Change()
      txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
End Sub
Private Sub gLista_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If UCase(Column.FieldName) = "TOTALGASTOVENTA" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub


Private Sub gdetalle_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError As String
Dim rsRpt As New ADODB.Recordset
On Error GoTo Err

    Select Case Column.Index
        Case GDetalle.Columns.ColumnByFieldName("ruc").Index
            If GDetalle.Columns.ColumnByFieldName("ruc").ReadOnly = False Then
                frmAyudaDocumentos_Gastos.MostrarForm txtCod_Vendedor.Text, rsRpt, Format(DtpFecha.Value, "dd/mm/yyyy"), StrMsgError
                If StrMsgError <> "" Then GoTo Err
                    procesaDocumentos rsRpt, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
            End If
    End Select

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
Private Sub gdetalle_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If Action = daInsert Then
        GDetalle.Columns.ColumnByFieldName("item").Value = GDetalle.Count
        GDetalle.Columns.ColumnByFieldName("idCliente").Value = ""
        GDetalle.Columns.ColumnByFieldName("idMoneda").Value = ""
        GDetalle.Columns.ColumnByFieldName("valSaldoOriginal").Value = 0
        GDetalle.Columns.ColumnByFieldName("valImporte").Value = 0
        GDetalle.Columns.ColumnByFieldName("ruc").Value = ""
        GDetalle.Columns.ColumnByFieldName("NroDocumento").Value = ""
        GDetalle.Columns.ColumnByFieldName("NroComprobante").Value = ""
        GDetalle.Columns.ColumnByFieldName("valSaldoInicial").Value = 0
        GDetalle.Dataset.Post
    End If
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
Dim rsg As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim strAno As String
On Error GoTo Err
     
    
    '********FORMATO GRILLA DETALLE
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idCliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idMoneda", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "valImporte", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "ruc", adVarChar, 15, adFldIsNullable
    rsg.Fields.Append "NroDocumento", adVarChar, 15, adFldIsNullable
    
    
    rsg.Open
    rsg.AddNew
    rsg.Fields("Item") = 1
    rsg.Fields("idCliente") = ""
    rsg.Fields("idMoneda") = ""
    rsg.Fields("valImporte") = 0
    rsg.Fields("ruc") = ""
    rsg.Fields("NroDocumento") = ""
    
    
    Set GDetalle.DataSource = Nothing
    mostrarDatosGridSQL GDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    GDetalle.Columns.FocusedIndex = GDetalle.Columns.ColumnByFieldName("ruc").ColIndex
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub procesaDocumentos(ByVal rsDoc As ADODB.Recordset, ByRef StrMsgError As String)
Dim i As Integer
Dim rsDetalle As New ADODB.Recordset
On Error GoTo Err

    rsDetalle.Fields.Append "Item", adInteger, , adFldRowID
    rsDetalle.Fields.Append "ruc", adVarChar, 11, adFldIsNullable
    rsDetalle.Fields.Append "GlsCliente", adVarChar, 200, adFldIsNullable
    rsDetalle.Fields.Append "idCliente", adVarChar, 8, adFldIsNullable
    rsDetalle.Fields.Append "NroDocumento", adVarChar, 25, adFldIsNullable
    rsDetalle.Fields.Append "idMoneda", adVarChar, 12, adFldIsNullable
    rsDetalle.Fields.Append "ValImporte", adDouble, 14, adFldIsNullable
    rsDetalle.Fields.Append "Girar", adVarChar, 14, adFldIsNullable
    rsDetalle.Fields.Append "TC", adDouble, 14, adFldIsNullable
    rsDetalle.Open
    
    GDetalle.Dataset.First
    Do While Not GDetalle.Dataset.EOF
       If Trim(GDetalle.Columns.ColumnByFieldName("ruc").Value) <> "" Then
            i = i + 1
            rsDetalle.AddNew
            rsDetalle.Fields("Item") = i
            rsDetalle.Fields("ruc") = Trim(GDetalle.Columns.ColumnByFieldName("ruc").Value)
            rsDetalle.Fields("NroDocumento") = Trim(GDetalle.Columns.ColumnByFieldName("NroDocumento").Value)
            rsDetalle.Fields("idMoneda") = Trim(GDetalle.Columns.ColumnByFieldName("idMoneda").Value)
            rsDetalle.Fields("idCliente") = Trim(GDetalle.Columns.ColumnByFieldName("idCliente").Value)
            rsDetalle.Fields("GlsCliente") = Trim(GDetalle.Columns.ColumnByFieldName("GlsCliente").Value)
            rsDetalle.Fields("valImporte") = Trim(GDetalle.Columns.ColumnByFieldName("valImporte").Value)
        End If
        GDetalle.Dataset.Next
    Loop
    
    rsDoc.MoveFirst
    Do While Not rsDoc.EOF
        
        If Trim("" & rsDoc("Girar").Value) = "S" Then
            i = i + 1
            rsDetalle.AddNew
            rsDetalle.Fields("Item") = i
            rsDetalle.Fields("ruc") = Trim(rsDoc("RUC") & "")
            rsDetalle.Fields("NroDocumento") = Trim(rsDoc("NroDocumento") & "")
            If Trim("" & rsDoc("idMoneda")) = "US$." Then
                rsDetalle.Fields("valImporte") = Format(Format(rsDoc("ValTotalDoc").Value, "0.00") * rsDoc("TC").Value, "0.00")
            Else
                rsDetalle.Fields("valImporte") = Trim("" & rsDoc("ValTotalDoc"))
            End If
            rsDetalle.Fields("idMoneda") = Trim(rsDoc("idMoneda") & "")
            rsDetalle.Fields("idCliente") = Trim(rsDoc("idCliente") & "")
            rsDetalle.Fields("GlsCliente") = traerCampo("Personas", "GlsPersona", "idPersona", Trim("" & rsDoc("idCliente")), False)
        End If
        rsDoc.MoveNext
        
    Loop
    
    Set rs_planilla_det = rsDetalle
    mostrarDatosGridSQL GDetalle, rsDetalle, StrMsgError
    If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    'Resume
End Sub
Private Sub EjecutaDetalleGasto(ByRef StrMsgError As String)
 Dim strCampos, strValores, strTipoDato, strCampo As String

    If TypeName(GDetalle) <> "Nothing" Then
        Cn.Execute "DELETE FROM gastosventadet WHERE idEmpresa = '" & glsEmpresa & "' AND   idCodGastoVenta = '" & txtCod_GastoVenta.Text & "'"
        
        GDetalle.Dataset.First
        Do While Not GDetalle.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To GDetalle.Columns.Count - 1
                    If UCase(left(GDetalle.Columns(i).ObjectName, 1)) = "W" Then
                        strTipoDato = Mid(GDetalle.Columns(i).ObjectName, 2, 1)
                        strCampo = Mid(GDetalle.Columns(i).ObjectName, 3)
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val("" & GDetalle.Columns(i).Value) & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(GDetalle.Columns(i).Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(GDetalle.Columns(i).Value, "yyyy-mm-dd") & "',"
                        End Select
                    End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO gastosventadet (" & strCampos & ",idCodGastoVenta,idEmpresa) VALUES(" & strValores & ",'" & txtCod_GastoVenta.Text & "','" & glsEmpresa & "')"
                    
            Cn.Execute csql
            
            GDetalle.Dataset.Next
        Loop
    End If
End Sub
Private Sub gdetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
    
    If KeyCode = 46 Then
        If GDetalle.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                                               
                If GDetalle.Count = 1 Then
                    GDetalle.Dataset.Edit
                    GDetalle.Columns.ColumnByFieldName("item").Value = GDetalle.Count
                    GDetalle.Columns.ColumnByFieldName("ruc").Value = ""
                    GDetalle.Columns.ColumnByFieldName("GlsCliente").Value = ""
                    GDetalle.Columns.ColumnByFieldName("idCliente").Value = 0
                    GDetalle.Columns.ColumnByFieldName("NroDocumento").Value = ""
                    GDetalle.Columns.ColumnByFieldName("idMoneda").Value = ""
                    GDetalle.Columns.ColumnByFieldName("ValImporte").Value = 0
                    GDetalle.Columns.ColumnByFieldName("TC").Value = 0
                    GDetalle.Dataset.Post
                Else
                    GDetalle.Dataset.Delete
                    GDetalle.Dataset.First
                    Do While Not GDetalle.Dataset.EOF
                        i = i + 1
                        GDetalle.Dataset.Edit
                        GDetalle.Columns.ColumnByFieldName("Item").Value = i
                        GDetalle.Dataset.Post
                        GDetalle.Dataset.Next
                    Loop
                    If GDetalle.Dataset.State = dsEdit Or GDetalle.Dataset.State = dsInsert Then
                        GDetalle.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    If KeyCode = 13 Then
        If GDetalle.Dataset.State = dsEdit Or GDetalle.Dataset.State = dsInsert Then
              GDetalle.Dataset.Post
        End If
    End If

End Sub





