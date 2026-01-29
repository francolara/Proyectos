VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantCombos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Fórmulas"
   ClientHeight    =   8490
   ClientLeft      =   2685
   ClientTop       =   1650
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   10305
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantCombos.frx":3518
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Excel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7785
      Left            =   45
      TabIndex        =   10
      Top             =   645
      Width           =   12240
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   150
         TabIndex        =   11
         Top             =   180
         Width           =   11895
         Begin VB.CommandButton CmdActualizaPrecios 
            Caption         =   "Actualizar Precios"
            Height          =   330
            Left            =   9720
            TabIndex        =   18
            Top             =   1170
            Width           =   1950
         End
         Begin VB.CommandButton cmbAyudaAlmacen 
            Height          =   315
            Left            =   11310
            Picture         =   "frmMantCombos.frx":396A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   740
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   0
            Left            =   11310
            Picture         =   "frmMantCombos.frx":3CF4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   370
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Almacen 
            Height          =   315
            Left            =   1230
            TabIndex        =   3
            Tag             =   "TidAlmacen"
            Top             =   735
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
            Container       =   "frmMantCombos.frx":407E
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Almacen 
            Height          =   315
            Left            =   2175
            TabIndex        =   14
            Top             =   735
            Width           =   9120
            _ExtentX        =   16087
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
            Container       =   "frmMantCombos.frx":409A
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Combo 
            Height          =   315
            Index           =   0
            Left            =   1230
            TabIndex        =   2
            Tag             =   "TidNivelPred"
            Top             =   375
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
            Container       =   "frmMantCombos.frx":40B6
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Combo 
            Height          =   315
            Index           =   0
            Left            =   2175
            TabIndex        =   15
            Top             =   375
            Width           =   9120
            _ExtentX        =   16087
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
            Container       =   "frmMantCombos.frx":40D2
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_Almacen 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
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
            Left            =   270
            TabIndex        =   17
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fórmula"
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
            Index           =   0
            Left            =   270
            TabIndex        =   16
            Top             =   435
            Width           =   570
         End
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   11610
         Top             =   -90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DXDBGRIDLibCtl.dxDBGrid GDetalle 
         Height          =   5655
         Left            =   150
         OleObjectBlob   =   "frmMantCombos.frx":40EE
         TabIndex        =   4
         Top             =   1980
         Width           =   11925
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   7770
      Left            =   45
      TabIndex        =   6
      Top             =   645
      Width           =   12240
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   135
         TabIndex        =   7
         Top             =   150
         Width           =   11985
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1095
            TabIndex        =   0
            Top             =   255
            Width           =   10770
            _ExtentX        =   18997
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
            Container       =   "frmMantCombos.frx":6ED1
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
            Left            =   225
            TabIndex        =   8
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3465
         Left            =   135
         OleObjectBlob   =   "frmMantCombos.frx":6EED
         TabIndex        =   1
         Top             =   915
         Width           =   12000
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   3135
         Left            =   135
         OleObjectBlob   =   "frmMantCombos.frx":9D70
         TabIndex        =   5
         Top             =   4530
         Width           =   12000
      End
   End
End
Attribute VB_Name = "frmMantCombos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numPesos As Integer
Dim indIngresoImagen As Boolean
Dim indCalculando As Boolean
Dim strTipoDoc As String
Dim objDocVentas As New clsDocVentas
Dim prodRef As String
Dim intBoton As Integer
Dim indCargando As Boolean
Dim indNuevoDoc As Boolean
Dim stado As String
Dim csq As String
Dim TipoDoc As String
Private indInserta  As Boolean
Private indInsertaDocRef  As Boolean
Private strEstDocVentas As String
Private indGeneraVale As Boolean
Private strGlsTipoDoc As String
Dim intRegMax As Integer
Dim idDocVentasImpor_Temp   As String

Private Sub cmbAyudaAlmacen_Click()
    
    mostrarAyuda "ALMACEN", txtCod_Almacen, txtGls_Almacen
    If txtCod_Almacen.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub CmbAyudaNivel_Click(Index As Integer)
    
    mostrarAyuda "PRODUCTOS", txtCod_Combo(Index), txtGls_Combo(Index), " AND idTipoProducto = '06004'"

End Sub

Private Sub CmdActualizaPrecios_Click()
On Error GoTo Err
Dim StrMsgError                 As String
Dim X                           As Double
Dim CArrays(3)                  As String

    gDetalle.Dataset.First
    Do While Not gDetalle.Dataset.EOF
        
        If Len(Trim(gDetalle.Columns.ColumnByFieldName("IdProducto").Value)) > 0 Then
            
            traerCampos "PreciosVenta V Inner Join Productos P On V.IdEmpresa = P.IdEmpresa And V.IdProducto = P.IdProducto", "P.GlsProducto,P.IdUMVenta,V.VVUnit", "V.IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, 3, CArrays, False, "V.IdEmpresa = '" & glsEmpresa & "' And V.IdLista = '" & glsListaVentas & "'"
            
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = CArrays(0)
            gDetalle.Columns.ColumnByFieldName("IdUM").Value = CArrays(1)
            
            If Val(CArrays(2)) > 0 Then
                
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = Val(CArrays(2))
                gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = Val(CArrays(2))
                
            End If
            
            gDetalle.Dataset.Post
            
        End If
        
        gDetalle.Dataset.Next
        
    Loop
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
    
    indInserta = False
    indInsertaDocRef = False
    indNuevoDoc = True
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gListaDetalle, False, False, False, False
    ConfGrid gDetalle, True, False, False, False
    
    listaProducto StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 6
    nuevo
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo               As String
Dim strMsg                  As String, cbusca As String
Dim strUM                   As String
Dim strMoneda               As String
Dim NTotPrecio              As Double
Dim indTrans                As Boolean

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    eliminaNulosGrilla
    
    If gDetalle.Count >= 1 Then
        If gDetalle.Count = 1 And (gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" Or gDetalle.Columns.ColumnByFieldName("Cantidad").Value <= 0) Then
            StrMsgError = "Falta Ingresar Detalle"
            GoTo Err
        End If
    End If

    strUM = traerCampo("productos", "idUMVenta", "idProducto", txtCod_Combo(0).Text, True)
    strMoneda = traerCampo("productos", "idMoneda", "idProducto", txtCod_Combo(0).Text, True)
    
    NTotPrecio = 0
    
    If intBoton = 1 Then '--- Graba
        
        Cn.BeginTrans
        indTrans = True
        
        csql = "Insert Into ComboCab(IdComboCab,GlsCombo,IdUM,FecEmision,IdUsuario,IdMoneda,IdEmpresa,IdSucursal,IdAlmacen)Values(" & _
               "'" & txtCod_Combo(0).Text & "', '" & txtGls_Combo(0).Text & "', '" & strUM & "', sysdate(), '" & glsUser & "', '" & strMoneda & "'," & _
               "'" & glsEmpresa & "','" & glsSucursal & "', '" & txtCod_Almacen.Text & "')"
        Cn.Execute csql
        
        gDetalle.Dataset.First
        Do While Not gDetalle.Dataset.EOF
            If gDetalle.Columns.ColumnByFieldName("idProducto").Value <> "" And Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) <> 0 Then
                csql = "Insert Into ComboDet(IdComboCab,IdProducto,GlsProducto,Cantidad,Item,IdEmpresa,IdSucursal,IdTipoProducto,VVUnit,PorcDcto," & _
                       "TotalVVNeto)Values(" & _
                       "'" & txtCod_Combo(0).Text & "', '" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value & "'," & _
                       "'" & gDetalle.Columns.ColumnByFieldName("GlsProducto").Value & "'," & _
                       "" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value & "," & _
                       "" & gDetalle.Columns.ColumnByFieldName("Item").Value & ",'" & glsEmpresa & "','" & glsSucursal & "'," & _
                       "'" & gDetalle.Columns.ColumnByFieldName("IdTipoProducto").Value & "'," & _
                       "" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "," & _
                       "" & gDetalle.Columns.ColumnByFieldName("PorcDcto").Value & "," & _
                       "" & gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value & ")"
                Cn.Execute csql
                
                NTotPrecio = NTotPrecio + Val("" & gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
                
            End If
            gDetalle.Dataset.Next
        Loop
        
        csql = "Update ComboCab " & _
               "Set TotalValorVenta = " & NTotPrecio & " " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdComboCab = '" & txtCod_Combo(0).Text & "'"
        
        Cn.Execute csql
        
        Cn.CommitTrans
        indTrans = False
        
        strMsg = "Grabo"
        intBoton = 3
    
    Else '--- Modifica
        
        Cn.BeginTrans
        indTrans = True
        
        csql = "delete from combocab where idComboCab = '" & txtCod_Combo(0).Text & "' and idEmpresa = '" & glsEmpresa & "' and idSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
        csql = "INSERT INTO combocab(idComboCab, glsCombo, idUM, fecEmision, idUsuario, idMoneda, IdEmpresa, idSucursal,IdAlmacen)VALUES(" & _
                "'" & txtCod_Combo(0).Text & "', '" & txtGls_Combo(0).Text & "', '" & strUM & "', sysdate(), '" & glsUser & "', '" & strMoneda & "'," & _
                "'" & glsEmpresa & "','" & glsSucursal & "', '" & txtCod_Almacen.Text & "')"
        Cn.Execute csql
        
        csql = "delete from combodet where idComboCab = '" & txtCod_Combo(0).Text & "' and idEmpresa = '" & glsEmpresa & "' and idSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
        gDetalle.Dataset.First
        Do While Not gDetalle.Dataset.EOF
            If gDetalle.Columns.ColumnByFieldName("idProducto").Value <> "" And Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) <> 0 Then
                csql = "Insert Into ComboDet(IdComboCab,IdProducto,GlsProducto,Cantidad,Item,IdEmpresa,IdSucursal,VVUnit,PorcDcto," & _
                       "TotalVVNeto)Values(" & _
                       "'" & txtCod_Combo(0).Text & "', '" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value & "'," & _
                       "'" & gDetalle.Columns.ColumnByFieldName("GlsProducto").Value & "'," & _
                       "" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value & "," & _
                       "" & gDetalle.Columns.ColumnByFieldName("Item").Value & ",'" & glsEmpresa & "','" & glsSucursal & "'," & _
                       "" & _
                       "" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "," & _
                       "" & gDetalle.Columns.ColumnByFieldName("PorcDcto").Value & "," & _
                       "" & gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value & ")"
                Cn.Execute csql
                
                NTotPrecio = NTotPrecio + Val("" & gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
                
            End If
            gDetalle.Dataset.Next
        Loop
        
        csql = "Update ComboCab " & _
               "Set TotalValorVenta = " & NTotPrecio & " " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdComboCab = '" & txtCod_Combo(0).Text & "'"
        
        Cn.Execute csql
        
        Cn.CommitTrans
        indTrans = False
        
        strMsg = "Modifico"
        
    End If

    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaProducto StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If indTrans Then Cn.RollbackTrans: indTrans = False
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gDetalle.Count >= 1 Then
            gDetalle.Dataset.First
            indEntro = False
            Do While Not gDetalle.Dataset.EOF
                If Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value) = "" Or gDetalle.Columns.ColumnByFieldName("Cantidad").Value <= 0 Then
                    gDetalle.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gDetalle.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gDetalle.Count >= 1 Then
        gDetalle.Dataset.First
        i = 0
        Do While Not gDetalle.Dataset.EOF
            i = i + 1
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("item").Value = i
            If gDetalle.Dataset.State = dsEdit Then gDetalle.Dataset.Post
            gDetalle.Dataset.Next
        Loop
    Else
        indInserta = True
        gDetalle.Dataset.Append
        indInserta = False
    End If

End Sub

Private Sub nuevo()
Dim rsg As New ADODB.Recordset
Dim StrMsgError As String

    limpiaForm Me
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorcDcto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Open
    
    rsg.AddNew
    rsg.Fields("Item") = 1
    rsg.Fields("idProducto") = ""
    rsg.Fields("GlsProducto") = ""
    rsg.Fields("idMarca") = ""
    rsg.Fields("idUM") = ""
    rsg.Fields("Cantidad") = 0
    rsg.Fields("VVUnit") = 0
    rsg.Fields("PorcDcto") = 0
    rsg.Fields("TotalVVNeto") = 0

    Set gDetalle.DataSource = Nothing
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").Index
    listaProducto StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gdetalle_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)

    If Action = daInsert Then
        gDetalle.Columns.ColumnByFieldName("item").Value = gDetalle.Count
        gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
        gDetalle.Dataset.Post
    End If
    
End Sub

Private Sub gdetalle_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If intRegMax = 0 Or gDetalle.Count < intRegMax Then
            gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").ColIndex
        Else
            Allow = False
        End If
    End If

End Sub

Private Sub gDetalle_OnChangeColumn(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal OldColumn As DXDBGRIDLibCtl.IdxGridColumn, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn)

    If gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("GlsProducto").Index Or gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("PVUnit").Index Or gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("VVUnit").Index Then
        If gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = "06002" Then 'Servicios
            gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = False
        Else
            gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = True
        End If
    End If
    
End Sub

Private Sub gDetalle_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String

    If traerCampo("Productos V", "V.IdTipoProducto", "V.IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, True) = "06002" Then
        gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
    Else
        gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = True
        If Val("" & traerCampo("PreciosVenta V", "V.VVUnit", "V.IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, True, "V.IdLista = '" & glsListaVentas & "'")) > 0 Then
            gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = True
        Else
            gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
        End If
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gdetalle_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
Dim StrCod                      As String
Dim StrDes                      As String
Dim strTipoProd                 As String
Dim intFila                     As Integer
Dim indPedido                   As Boolean
Dim rspa                        As New ADODB.Recordset
    
    intFila = Node.Index + 1
    Select Case Column.Index
        Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
            StrCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
            StrDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
            indPedido = False
            
            FrmAyudaProdOC.ExecuteReturnTextAlm txtCod_Almacen.Text, rspa, StrCod, StrDes, "", glsValidaStock, "", False, True, indPedido, StrMsgError
            
            gDetalle.SetFocus
            If rspa.RecordCount <> 0 Then
                mostrarProdImp rspa, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            gDetalle.Dataset.RecNo = intFila
            
    End Select
    gDetalle.Dataset.RecNo = intFila
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarProdImp(ByVal rsdd As ADODB.Recordset, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim strSerieDocVentas As String
Dim dblTC  As Double
Dim strCodFabri As String
Dim strCodMar As String
Dim strDesMar As String
Dim intAfecto As Integer
Dim strTipoProd As String
Dim strMoneda As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblVVUnit  As Double
Dim dblIGVUnit  As Double
Dim dblPVUnit  As Double
Dim dblFactor  As Double
Dim intFila As Integer
Dim i As Integer
Dim indExisteDocRef As Boolean
Dim primero As Boolean
Dim strInserta As Boolean

    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        strInserta = True
        If strInserta = True Then
        
            If primero = True Then
                primero = False
            Else
                gDetalle.Dataset.Insert
            End If
        
            gDetalle.SetFocus
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
            gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
            gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 1
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = Val("" & traerCampo("PreciosVenta", "VVUnit", "IdProducto", "" & rsdd.Fields("idProducto"), True, "IdLista = '" & glsListaVentas & "'"))
            gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = 0
            gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = Val(gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "")
            gDetalle.Dataset.Post
            gDetalle.Dataset.RecNo = intFila
            gDetalle.Dataset.Edit
                            
            gDetalle.Dataset.Post
            If "" & rsdd.Fields("idProducto") <> "" Then
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
            End If
        End If
        rsdd.MoveNext
    Loop

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gDetalle_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim RsP As New ADODB.Recordset
Dim StrMsgError As String
Dim StrCod As String
Dim StrDes As String
Dim dblTC  As Double
Dim strCodFabri As String
Dim strCodMar As String
Dim strDesMar As String
Dim intAfecto As Integer
Dim strTipoProd As String
Dim strMoneda As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblVVUnit  As Double
Dim dblIGVUnit  As Double
Dim dblPVUnit  As Double
Dim dblFactor  As Double
Dim dblDcto As Double
Dim dblTotalBruto As Double
Dim indEvaluacion As Integer
Dim strCodUsuarioAutorizacion As String
Dim i As Integer
Dim sw_dcto     As Boolean
Dim ndcto       As Double
Dim X                           As Double
Dim CArrays(3)                  As String

    If gDetalle.Dataset.Modified = False Then Exit Sub
    
    Select Case gDetalle.Columns.FocusedColumn.Index
        Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
            If Len(Trim(gDetalle.Columns.ColumnByFieldName("IdProducto").Value)) > 0 Then
                traerCampos "PreciosVenta V Inner Join Productos P On V.IdEmpresa = P.IdEmpresa And V.IdProducto = P.IdProducto", "P.GlsProducto,P.IdUMVenta,V.VVUnit", "V.IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, 3, CArrays, False, "V.IdEmpresa = '" & glsEmpresa & "' And V.IdLista = '" & glsListaVentas & "'"
                gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = CArrays(0)
                gDetalle.Columns.ColumnByFieldName("IdUM").Value = CArrays(1)
                gDetalle.Columns.ColumnByFieldName("Cantidad").Value = "1"
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = Val(CArrays(2))
                gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = "0"
                gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = Val(CArrays(2))
                        
            Else
                gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                gDetalle.Columns.ColumnByFieldName("IdUM").Value = ""
                gDetalle.Columns.ColumnByFieldName("Cantidad").Value = "0"
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = "0"
                gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = "0"
                gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = "0"
            End If
        
        Case gDetalle.Columns.ColumnByFieldName("Cantidad").Index, gDetalle.Columns.ColumnByFieldName("VVUnit").Index, gDetalle.Columns.ColumnByFieldName("PorcDcto").Index
            gDetalle.Columns.ColumnByFieldName("Cantidad").Value = Val(gDetalle.Columns.ColumnByFieldName("Cantidad").Value & "")
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = Val(gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "")
            gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = Val(gDetalle.Columns.ColumnByFieldName("PorcDcto").Value & "")
            
            If Val(gDetalle.Columns.ColumnByFieldName("PorcDcto").Value & "") > Val("" & traerCampo("PreciosVenta", "MaxDcto", "IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, True, "IdLista = '" & glsListaVentas & "'")) Then
                gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = "0"
                StrMsgError = "El Descuento ingresado es mayor al Máximo descuento asignado a producto,Verifique."
            End If
            
            If gDetalle.Columns.ColumnByFieldName("PorcDcto").Value > "0" Then
                gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = (Val(gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "") - (Val(gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "") * (Val(gDetalle.Columns.ColumnByFieldName("PorcDcto").Value & "") / 100))) * Val(gDetalle.Columns.ColumnByFieldName("Cantidad").Value & "")
            Else
                gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = Val(gDetalle.Columns.ColumnByFieldName("VVUnit").Value & "") * Val(gDetalle.Columns.ColumnByFieldName("Cantidad").Value & "")
            End If
            If StrMsgError <> "" Then GoTo Err
    End Select
    If RsP.State = 1 Then RsP.Close: Set RsP = Nothing
    
    Exit Sub
    
Err:
    If RsP.State = 1 Then RsP.Close: Set RsP = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gdetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gDetalle.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gDetalle.Count = 1 Then
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("item").Value = 1
                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PorcDcto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
                    gDetalle.Dataset.Post
                
                Else
                    gDetalle.Dataset.Delete
                    gDetalle.Dataset.First
                    Do While Not gDetalle.Dataset.EOF
                        i = i + 1
                        gDetalle.Dataset.Edit
                        gDetalle.Columns.ColumnByFieldName("Item").Value = i
                        gDetalle.Dataset.Post
                        gDetalle.Dataset.Next
                    Loop
                    If gDetalle.Dataset.State = dsEdit Or gDetalle.Dataset.State = dsInsert Then
                        gDetalle.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gDetalle.Dataset.State = dsEdit Or gDetalle.Dataset.State = dsInsert Then
            gDetalle.Dataset.Post
        End If
    End If

End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    listaDetalle

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarProducto gLista.Columns.ColumnByName("idComboCab").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    intBoton = 3
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnReloadGroupList()
    
    gLista.m.FullExpand

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError     As String
Dim strCodUltProd   As String
Dim CIdCombo        As String

    Select Case Button.Index
        Case 1 'Nuevo
            intBoton = Button.Index
            nuevo
            indCalculando = False
            If StrMsgError <> "" Then GoTo Err
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            intBoton = 3
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4, 6 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 7
            gLista.m.ExportToXLS App.Path & "\Temporales\Productos.xls"
            ShellEx App.Path & "\Temporales\Productos.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 8 'Importar
            CIdCombo = txtCod_Combo(0).Text
            FrmImportaCombos.MostrarForm StrMsgError, CIdCombo
            If StrMsgError <> "" Then GoTo Err
            
            If Len(Trim(CIdCombo)) > 0 Then
                MostrarComboImportado StrMsgError, CIdCombo
                If StrMsgError <> "" Then GoTo Err
            End If
        Case 9 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
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
            Toolbar1.Buttons(6).Visible = indHabilitar 'lista
            Toolbar1.Buttons(7).Visible = indHabilitar 'excel
            Toolbar1.Buttons(8).Visible = Not indHabilitar 'Importar
        Case 4, 6 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = False
    End Select

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaProducto StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaProducto(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim intNumNiveles As Integer
Dim strTabla As String
Dim strWhere As String
Dim strCampos As String
Dim strTablas As String
Dim strTablaAnt As String
Dim i As Integer

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (glsCombo LIKE '%" & strCond & "%' or idComboCab LIKE '%" & strCond & "%') "
    End If

    csql = "SELECT idComboCab,glsCombo,DATE_FORMAT(FecEmision,GET_FORMAT(DATE, 'EUR')) as FecEmision,Format(TotalValorVenta,2) AS TotalValorVenta, idUM,idMoneda, idUsuario " & _
            "FROM combocab " & _
            "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'"
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY idComboCab, FecEmision"

    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idComboCab"
    End With

    listaDetalle
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaDetalle()

    csql = "SELECT item, idProducto, GlsProducto, idUM, Format(Cantidad,2) AS Cantidad, Format(PVUnit,2) AS PVUnit " & _
           "FROM combodet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idComboCab = '" & gLista.Columns.ColumnByFieldName("idComboCab").Value & "'"
    
    With gListaDetalle
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With

End Sub

Private Sub mostrarProducto(strCodProd As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsu As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim strIndImagen As String
Dim i As Integer

    txtCod_Combo(0).Text = strCodProd
    txtCod_Almacen.Text = traerCampo("ComboCab", "IdAlmacen", "idComboCab", strCodProd, True)
    
    csql = "SELECT * " & _
           "FROM combodet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idcombocab = '" & strCodProd & "' ORDER BY ITEM"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    rsg.Fields.Append "Item", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdProducto", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 240, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdComboCab", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorcDcto", adInteger, , adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("Cantidad") = 0
        rsg.Fields("IdComboCab") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("PorcDcto") = 0
        rsg.Fields("TotalVVNeto") = 0
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idProducto") = "" & rst.Fields("idProducto")
            rsg.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsg.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsg.Fields("IdComboCab") = 0
            rsg.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            rsg.Fields("PorcDcto") = "" & rst.Fields("PorcDcto")
            rsg.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Almacen_Change()
    
    txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)
    
End Sub

Private Sub txtCod_Combo_Change(Index As Integer)
    
    txtGls_Combo(0).Text = traerCampo("productos", "glsProducto", "idProducto", txtCod_Combo(0).Text, True)

End Sub

Private Sub eliminar(ByRef StrMsgError As String)

    If MsgBox("¿Seguro de eliminar el documento?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    csql = "delete from combocab where idComboCab = '" & txtCod_Combo(0).Text & "' and idEmpresa = '" & glsEmpresa & "' and idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    csql = "delete from combodet where idComboCab = '" & txtCod_Combo(0).Text & "' and idEmpresa = '" & glsEmpresa & "' and idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    csql = "delete from preciosventa where idProducto = '" & txtCod_Combo(0).Text & "' and idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title

End Sub

Private Sub muestraColumnasDetalle()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim pCtrl As Object
Dim strSerie As String

    strSerie = "001"
    csql = "SELECT GlsObj,etiqueta,numCol,ancho,Tipodato,Decimales  FROM objdocventas " & _
            "where idEmpresa = '" & glsEmpresa & "' and idDocumento = '99' and idserie = '" & strSerie & "' and tipoObj = 'D' and indVisible = 'V' "
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    Do While Not rst.EOF
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Caption = rst.Fields("etiqueta") & ""
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").ColIndex = Val(rst.Fields("numCol") & "")
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Width = Val(rst.Fields("ancho") & "")
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Visible = True
        If (rst.Fields("Tipodato") & "") = "N" Then
            gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").DecimalPlaces = Val(rst.Fields("Decimales") & "")
        End If
        rst.MoveNext
    Loop
    
    Exit Sub

Err:
    MsgBox Err.Description
End Sub

Private Function DatosProducto(strCodProd As String, ByRef strCodFabri As String, ByRef strCodMar As String, ByRef strGlsMarca As String, ByRef intAfecto As Integer, ByRef strTipoProd As String) As Boolean
Dim rst As New ADODB.Recordset

    csql = "SELECT p.idFabricante,p.idMarca,m.GlsMarca,p.AfectoIGV,p.idTipoProducto " & _
            "FROM productos p LEFT JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca AND m.idEmpresa = '" & glsEmpresa & "' " & _
            "WHERE p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idProducto = '" & strCodProd & "'"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        DatosProducto = True
        strCodFabri = "" & rst.Fields("idFabricante")
        strCodMar = "" & rst.Fields("idMarca")
        strGlsMarca = "" & rst.Fields("GlsMarca")
        intAfecto = "" & rst.Fields("AfectoIGV")
        strTipoProd = "" & rst.Fields("idTipoProducto")
    Else
        DatosProducto = False
        strCodFabri = ""
        strCodMar = ""
        strGlsMarca = ""
        intAfecto = 1
        strTipoProd = ""
    End If
    rst.Close: Set rst = Nothing
    
End Function

Private Function DatosPrecio(ByVal strCodProd As String, ByVal strTipoProd As String, ByVal strCodUM As String, ByRef strGlsUM As String, ByRef dblVVUnit As Double, ByRef dblFactor As Double) As Boolean
Dim rst As New ADODB.Recordset

    If strTipoProd = "06002" Then
        csql = "SELECT '' as GlsUM,v.VVUnit,1 AS Factor " & _
                "FROM preciosventa v " & _
                "WHERE v.idEmpresa = '" & glsEmpresa & "' " & _
                "AND v.idLista = '08040001'" & _
                "AND v.idProducto = '" & strCodProd & "'"
    Else
        csql = "SELECT u.abreUM as GlsUM,v.VVUnit,r.Factor " & _
                "FROM presentaciones r,unidadMedida u,preciosventa v " & _
                "WHERE r.idUM = u.idUM " & _
                "AND r.idEmpresa = '" & glsEmpresa & "' " & _
                "AND r.idProducto = v.idProducto " & _
                "AND r.idUM = v.idUM " & _
                "AND v.idEmpresa = '" & glsEmpresa & "' " & _
                "AND v.idLista = '08040001'" & _
                "AND r.idProducto = '" & strCodProd & "' " & _
                "AND r.idUM = '" & strCodUM & "'"
    End If
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        DatosPrecio = True
        strGlsUM = "" & rst.Fields("GlsUM")
        dblVVUnit = "" & rst.Fields("VVUnit")
        dblFactor = "" & rst.Fields("Factor")
    Else
        DatosPrecio = False
        strGlsUM = ""
        dblVVUnit = 0
        dblFactor = 1
    End If
    rst.Close: Set rst = Nothing
    
End Function

Private Sub MostrarComboImportado(StrMsgError As String, PIdCombo As String)
On Error GoTo Err
Dim CSqlC                       As String
Dim rst                         As New ADODB.Recordset
Dim rsg                         As New ADODB.Recordset

    CSqlC = "SELECT * " & _
            "FROM combodet " & _
            "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idcombocab = '" & PIdCombo & "' ORDER BY ITEM"
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    rsg.Fields.Append "Item", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdProducto", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 240, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdComboCab", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorcDcto", adInteger, , adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("Cantidad") = 0
        rsg.Fields("IdComboCab") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("PorcDcto") = 0
        rsg.Fields("TotalVVNeto") = 0
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idProducto") = "" & rst.Fields("idProducto")
            rsg.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsg.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsg.Fields("IdComboCab") = 0
            rsg.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            rsg.Fields("PorcDcto") = "" & rst.Fields("PorcDcto")
            rsg.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
