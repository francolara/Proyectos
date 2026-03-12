VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmRptVentaGeneticaLiquida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Venta de Genética Líquida"
   ClientHeight    =   4380
   ClientLeft      =   5715
   ClientTop       =   1350
   ClientWidth     =   7500
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
   ScaleHeight     =   4380
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExcel 
      Caption         =   "&Excel"
      Height          =   495
      Left            =   3105
      TabIndex        =   26
      Top             =   3760
      Width           =   1395
   End
   Begin VB.CommandButton cmbAyudaMoneda 
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
      Left            =   6720
      Picture         =   "frmRptVentaGeneticaLiquida.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2090
      Width           =   435
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3765
      Width           =   1300
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3765
      Width           =   1300
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   "Agrupado por "
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   255
      TabIndex        =   5
      Top             =   2550
      Width           =   6870
      Begin VB.OptionButton optTipo 
         Caption         =   "Producto"
         Height          =   240
         Index           =   2
         Left            =   5025
         TabIndex        =   8
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   22
         Top             =   6600
         Width           =   2025
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   15360
         TabIndex        =   21
         Top             =   360
         Width           =   2025
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Cliente"
         Height          =   240
         Index           =   1
         Left            =   2985
         TabIndex        =   7
         Top             =   360
         Width           =   1185
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Mes"
         Height          =   240
         Index           =   0
         Left            =   885
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   7395
      Begin VB.CommandButton cmbAyudaProducto 
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
         Left            =   6705
         Picture         =   "frmRptVentaGeneticaLiquida.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   435
      End
      Begin VB.CommandButton cmbAyudaCliente 
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
         Left            =   6705
         Picture         =   "frmRptVentaGeneticaLiquida.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1305
         Width           =   435
      End
      Begin VB.Frame frame2 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Left            =   255
         TabIndex        =   12
         Top             =   270
         Width           =   6870
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1560
            TabIndex        =   0
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
            Format          =   106758145
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4605
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
            Format          =   106758145
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   14
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4140
            TabIndex        =   13
            Top             =   375
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1065
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1260
         Width           =   1050
         _ExtentX        =   1852
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
         Container       =   "frmRptVentaGeneticaLiquida.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   1290
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "frmRptVentaGeneticaLiquida.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   1065
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1665
         Width           =   1050
         _ExtentX        =   1852
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
         Container       =   "frmRptVentaGeneticaLiquida.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   2160
         TabIndex        =   19
         Top             =   1665
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "frmRptVentaGeneticaLiquida.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   2070
         Width           =   1050
         _ExtentX        =   1852
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
         Container       =   "frmRptVentaGeneticaLiquida.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2160
         TabIndex        =   23
         Top             =   2070
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "frmRptVentaGeneticaLiquida.frx":0B2A
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   2145
         Width           =   570
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   20
         Top             =   1770
         Width           =   645
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   17
         Top             =   1335
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRptVentaGeneticaLiquida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaCliente_Click()

    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()

    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

End Sub

Private Sub cmbAyudaProducto_Click()
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim fIni                    As String, Ffin As String
Dim StrMsgError             As String
Dim CGlsReporte             As String
Dim GlsForm                 As String
Dim ctipo                   As String
Dim strMoneda               As String

    Screen.MousePointer = 11
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
    
    If optTipo(0).Value Then
        CGlsReporte = "RptVentaGeneticaLiquidaxMes.rpt"
        ctipo = "Por Mes"
    ElseIf optTipo(1).Value Then
        CGlsReporte = "RptVentaGeneticaLiquidaxCliente.rpt"
        ctipo = "Por Cliente"
    Else
        CGlsReporte = "RptVentaGeneticaLiquidaxProducto.rpt"
        ctipo = "Por Producto"
    End If
    
    GlsForm = Me.Caption & " - Detallado " & ctipo
    mostrarReporte CGlsReporte, "parEmpresa|parFechaIni|parFechaFin|ParProducto|parCliente|parMoneda", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Cliente.Text) & "|" & strMoneda, GlsForm, StrMsgError
    
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdExcel_Click()
On Error GoTo Err
Dim StrMsgError                         As String
Dim xl                                  As Excel.Application
Dim wb                                  As Workbook
Dim NFil                                As Long
Dim CSqlC                               As String
Dim RsC                                 As New ADODB.Recordset
Dim CGrupoAux                           As String
Dim cGrupo                              As String
Dim NVenta                              As Double
Dim NIGV                                As Double
Dim NTotal                              As Double
Dim NVentaTot                           As Double
Dim NIgvTot                             As Double
Dim NTotalTot                           As Double
Dim NVentaS                             As Double
Dim NVentaD                             As Double
Dim NIGVS                               As Double
Dim NIGVD                               As Double
Dim NTotalS                             As Double
Dim NTotalD                             As Double
Dim NVentaTotS                          As Double
Dim NVentaTotD                          As Double
Dim NIGVTotS                            As Double
Dim NIGVTotD                            As Double
Dim NTotalTotS                          As Double
Dim NTotalTotD                          As Double

    On Error GoTo ExcelNoAbierto
    Set xl = GetObject(, "Excel.Application")
    GoTo YaEstabaAbierto
ExcelNoAbierto:
    Set xl = New Excel.Application
YaEstabaAbierto:
    On Error GoTo 0

On Error GoTo Err

    Set wb = xl.Workbooks.Open(App.Path & "\Temporales\VentasGeneticaLiquida.xlt")
    xl.Visible = False
    
    NFil = 6
    
    xl.Cells(1, 1).Value = traerCampo("Empresas", "GlsEmpresa", "IdEmpresa", glsEmpresa, False)
    
    If optTipo(0).Value Then
        
        xl.Cells(1, 7).Value = "REPORTE DE VENTA DE GENETICA LIQUIDA DETALLADO POR MES"
        cGrupo = "IdMes"
        
        xl.Cells(5, 2).Value = "Fecha"
        xl.Cells(5, 3).Value = "Documento"
        xl.Cells(5, 4).Value = "Cliente"
        xl.Cells(5, 5).Value = "Unidad"
        xl.Cells(5, 6).Value = "Descripción"
        xl.Cells(5, 7).Value = "Moneda"
        xl.Cells(5, 8).Value = "P. Unit."
        xl.Cells(5, 9).Value = "V. Venta"
        xl.Cells(5, 10).Value = "Igv"
        xl.Cells(5, 11).Value = "Total"
        
    ElseIf optTipo(1).Value Then
    
        xl.Cells(1, 7).Value = "REPORTE DE VENTA DE GENETICA LIQUIDA DETALLADO POR CLIENTE"
        cGrupo = "Cliente"
        
        xl.Cells(5, 2).Value = "Fecha"
        xl.Cells(5, 3).Value = "Documento"
        xl.Cells(5, 4).Value = "Unidad"
        xl.Cells(5, 5).Value = "Descripción"
        xl.Cells(5, 6).Value = "Moneda"
        xl.Cells(5, 7).Value = "P. Unit."
        xl.Cells(5, 8).Value = "V. Venta"
        xl.Cells(5, 9).Value = "Igv"
        xl.Cells(5, 10).Value = "Total"
        
    ElseIf optTipo(2).Value Then
    
        xl.Cells(1, 7).Value = "REPORTE DE VENTA DE GENETICA LIQUIDA DETALLADO POR PRODUCTO"
        cGrupo = "GlsProducto"
        
        xl.Cells(5, 2).Value = "Cliente y Tipo"
        xl.Cells(5, 3).Value = "Fecha"
        xl.Cells(5, 4).Value = "Documento"
        xl.Cells(5, 5).Value = "Unidad"
        xl.Cells(5, 6).Value = "Moneda"
        xl.Cells(5, 7).Value = "P. Unit."
        xl.Cells(5, 8).Value = "V. Venta"
        xl.Cells(5, 9).Value = "Igv"
        xl.Cells(5, 10).Value = "Total"
        
    End If
    
    xl.Cells(5, 2).Font.Bold = True
    xl.Cells(5, 3).Font.Bold = True
    xl.Cells(5, 4).Font.Bold = True
    xl.Cells(5, 5).Font.Bold = True
    xl.Cells(5, 6).Font.Bold = True
    xl.Cells(5, 7).Font.Bold = True
    xl.Cells(5, 8).Font.Bold = True
    xl.Cells(5, 9).Font.Bold = True
    xl.Cells(5, 10).Font.Bold = True
    xl.Cells(5, 11).Font.Bold = True
    
    xl.Cells(1, 7).Font.Bold = True
    
    xl.Cells(2, 7).Value = "DEL " & dtpfInicio.Value & " AL " & dtpFFinal.Value
    
    xl.Cells(3, 7).Value = "EXPRESADO EN " & txtGls_Moneda.Text
    xl.Cells(3, 7).HorizontalAlignment = xlCenter
    
    CSqlC = "Call spu_ListaVentaGeneticaLiquida('" & glsEmpresa & "','" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "','" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "','" & txtCod_Producto.Text & "','" & txtCod_Cliente.Text & "','" & txtCod_Moneda.Text & "')"
    RsC.Open CSqlC, Cn, adOpenKeyset, adLockReadOnly
    If Not RsC.EOF Then
        
        If optTipo(0).Value Then
            RsC.Sort = "IdMes,FechaEmision,Documento"
        ElseIf optTipo(1).Value Then
            RsC.Sort = "Cliente,FechaEmision,Documento"
        ElseIf optTipo(2).Value Then
            RsC.Sort = "GlsProducto,FechaEmision,Documento"
        End If
        
        NVentaTot = 0
        NIgvTot = 0
        NTotalTot = 0
        
        NVentaTotS = 0
        NVentaTotD = 0
        NIGVTotS = 0
        NIGVTotD = 0
        NTotalTotS = 0
        NTotalTotD = 0
        
        Do While Not RsC.EOF
            
            If CGrupoAux <> Trim("" & RsC.Fields(cGrupo)) Then
                
                CGrupoAux = Trim("" & RsC.Fields(cGrupo))
                
                If optTipo(0).Value Then
                    
                    Select Case Val(Trim("" & RsC.Fields(cGrupo)))
                        
                        Case 1: xl.Cells(NFil, 1).Value = "ENERO"
                        Case 2: xl.Cells(NFil, 1).Value = "FEBRERO"
                        Case 3: xl.Cells(NFil, 1).Value = "MARZO"
                        Case 4: xl.Cells(NFil, 1).Value = "ABRIL"
                        Case 5: xl.Cells(NFil, 1).Value = "MAYO"
                        Case 6: xl.Cells(NFil, 1).Value = "JUNIO"
                        Case 7: xl.Cells(NFil, 1).Value = "JULIO"
                        Case 8: xl.Cells(NFil, 1).Value = "AGOSTO"
                        Case 9: xl.Cells(NFil, 1).Value = "SEPTIEMBRE"
                        Case 10: xl.Cells(NFil, 1).Value = "OCTUBRE"
                        Case 11: xl.Cells(NFil, 1).Value = "NOVIEMBRE"
                        Case 12: xl.Cells(NFil, 1).Value = "NOVIEMBRE"
                          
                    End Select
                
                ElseIf optTipo(1).Value Then
                    
                    xl.Cells(NFil, 1).Value = Trim("" & RsC.Fields(cGrupo))
                    
                ElseIf optTipo(2).Value Then
                
                    xl.Cells(NFil, 1).Value = Trim("" & RsC.Fields(cGrupo))
                    
                End If
                
                xl.Cells(NFil, 1).Font.Bold = True
                
                If txtCod_Moneda.Text <> "" Then
                
                    NVenta = 0
                    NIGV = 0
                    NTotal = 0
                
                Else
                
                    NVentaS = 0
                    NVentaD = 0
                    NIGVS = 0
                    NIGVD = 0
                    NTotalS = 0
                    NTotalD = 0
                    
                End If
                    
                NFil = NFil + 1
                
            End If
            
            If optTipo(0).Value Then
            
                xl.Cells(NFil, 2).Value = Format("" & RsC.Fields("FechaEmision"), "dd/mm/yyyy")
                xl.Cells(NFil, 3).Value = "" & RsC.Fields("Documento")
                xl.Cells(NFil, 4).Value = "" & RsC.Fields("Cliente")
                xl.Cells(NFil, 5).Value = Val("" & RsC.Fields("Cantidad"))
                xl.Cells(NFil, 5).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 5).HorizontalAlignment = xlRight
                xl.Cells(NFil, 5).ColumnWidth = 10
                xl.Cells(NFil, 6).Value = "" & RsC.Fields("GlsProducto")
                xl.Cells(NFil, 6).ColumnWidth = 30
                xl.Cells(NFil, 6).HorizontalAlignment = xlLeft
                xl.Cells(NFil, 7).Value = "" & RsC.Fields("GlsMonedaDet")
                xl.Cells(NFil, 7).ColumnWidth = 20
                xl.Cells(NFil, 7).HorizontalAlignment = xlCenter
                xl.Cells(NFil, 8).Value = Val("" & RsC.Fields("PrecioUnit"))
                xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                xl.Cells(NFil, 9).Value = Val("" & RsC.Fields("ValorVenta"))
                xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                xl.Cells(NFil, 10).Value = Val("" & RsC.Fields("Igv"))
                xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                xl.Cells(NFil, 11).Value = Val("" & RsC.Fields("PrecioTotal"))
                xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 11).HorizontalAlignment = xlRight
                
            ElseIf optTipo(1).Value Then
                
                xl.Cells(NFil, 2).Value = Format("" & RsC.Fields("FechaEmision"), "dd/mm/yyyy")
                xl.Cells(NFil, 3).Value = "" & RsC.Fields("Documento")
                xl.Cells(NFil, 4).Value = Val("" & RsC.Fields("Cantidad"))
                xl.Cells(NFil, 4).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 4).HorizontalAlignment = xlRight
                xl.Cells(NFil, 4).ColumnWidth = 10
                xl.Cells(NFil, 5).Value = "" & RsC.Fields("GlsProducto")
                xl.Cells(NFil, 5).ColumnWidth = 30
                xl.Cells(NFil, 5).HorizontalAlignment = xlLeft
                xl.Cells(NFil, 6).Value = "" & RsC.Fields("GlsMonedaDet")
                xl.Cells(NFil, 6).ColumnWidth = 20
                xl.Cells(NFil, 6).HorizontalAlignment = xlCenter
                xl.Cells(NFil, 7).Value = Val("" & RsC.Fields("PrecioUnit"))
                xl.Cells(NFil, 7).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 7).HorizontalAlignment = xlRight
                xl.Cells(NFil, 8).Value = Val("" & RsC.Fields("ValorVenta"))
                xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                xl.Cells(NFil, 9).Value = Val("" & RsC.Fields("Igv"))
                xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                xl.Cells(NFil, 10).Value = Val("" & RsC.Fields("PrecioTotal"))
                xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                
            ElseIf optTipo(2).Value Then
                
                xl.Cells(NFil, 2).Value = "" & RsC.Fields("Cliente")
                xl.Cells(NFil, 2).HorizontalAlignment = xlLeft
                xl.Cells(NFil, 2).ColumnWidth = 45
                xl.Cells(NFil, 3).Value = Format("" & RsC.Fields("FechaEmision"), "dd/mm/yyyy")
                xl.Cells(NFil, 3).HorizontalAlignment = xlCenter
                xl.Cells(NFil, 4).Value = "" & RsC.Fields("Documento")
                xl.Cells(NFil, 4).ColumnWidth = 15
                xl.Cells(NFil, 5).Value = Val("" & RsC.Fields("Cantidad"))
                xl.Cells(NFil, 5).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 5).HorizontalAlignment = xlRight
                xl.Cells(NFil, 5).ColumnWidth = 10
                xl.Cells(NFil, 6).Value = "" & RsC.Fields("GlsMonedaDet")
                xl.Cells(NFil, 6).ColumnWidth = 20
                xl.Cells(NFil, 6).HorizontalAlignment = xlCenter
                xl.Cells(NFil, 7).Value = Val("" & RsC.Fields("PrecioUnit"))
                xl.Cells(NFil, 7).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 7).HorizontalAlignment = xlRight
                xl.Cells(NFil, 7).ColumnWidth = 20
                xl.Cells(NFil, 8).Value = Val("" & RsC.Fields("ValorVenta"))
                xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                xl.Cells(NFil, 9).Value = Val("" & RsC.Fields("Igv"))
                xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                xl.Cells(NFil, 10).Value = Val("" & RsC.Fields("PrecioTotal"))
                xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                
            End If
            
            If txtCod_Moneda.Text <> "" Then
            
                NVenta = NVenta + Val("" & RsC.Fields("ValorVenta"))
                NIGV = NIGV + Val("" & RsC.Fields("Igv"))
                NTotal = NTotal + Val("" & RsC.Fields("PrecioTotal"))
                
                NVentaTot = NVentaTot + Val("" & RsC.Fields("ValorVenta"))
                NIgvTot = NIgvTot + Val("" & RsC.Fields("Igv"))
                NTotalTot = NTotalTot + Val("" & RsC.Fields("PrecioTotal"))
            
            Else
                
                If Trim("" & RsC.Fields("GlsMonedaDet")) = "Soles" Then
                    NVentaS = NVentaS + Val("" & RsC.Fields("ValorVenta"))
                    NIGVS = NIGVS + Val("" & RsC.Fields("Igv"))
                    NTotalS = NTotalS + Val("" & RsC.Fields("PrecioTotal"))
                
                    NVentaTotS = NVentaTotS + Val("" & RsC.Fields("ValorVenta"))
                    NIGVTotS = NIGVTotS + Val("" & RsC.Fields("Igv"))
                    NTotalTotS = NTotalTotS + Val("" & RsC.Fields("PrecioTotal"))
                Else
                    NVentaD = NVentaD + Val("" & RsC.Fields("ValorVenta"))
                    NIGVD = NIGVD + Val("" & RsC.Fields("Igv"))
                    NTotalD = NTotalD + Val("" & RsC.Fields("PrecioTotal"))
                
                    NVentaTotD = NVentaTotD + Val("" & RsC.Fields("ValorVenta"))
                    NIGVTotD = NIGVTotD + Val("" & RsC.Fields("Igv"))
                    NTotalTotD = NTotalTotD + Val("" & RsC.Fields("PrecioTotal"))
                End If
                
            End If
            
            RsC.MoveNext
            
            NFil = NFil + 1
            
            If Not RsC.EOF Then
            
                If CGrupoAux <> Trim("" & RsC.Fields(cGrupo)) Then
                    
                    If optTipo(0).Value Then
                        
                        If txtCod_Moneda.Text <> "" Then
                            xl.Cells(NFil, 8).Value = "Sub Total"
                            xl.Cells(NFil, 8).Font.Bold = True
                            
                            xl.Cells(NFil, 9).Value = NVenta
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 10).Value = NIGV
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 11).Value = NTotal
                            xl.Cells(NFil, 11).Font.Bold = True
                            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
                        Else
                            xl.Cells(NFil, 8).Value = "Sub Total Soles"
                            xl.Cells(NFil, 8).Font.Bold = True
                            
                            xl.Cells(NFil, 9).Value = NVentaS
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 10).Value = NIGVS
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 11).Value = NTotalS
                            xl.Cells(NFil, 11).Font.Bold = True
                            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
                            
                            NFil = NFil + 1
                            
                            xl.Cells(NFil, 8).Value = "Sub Total Dolares"
                            xl.Cells(NFil, 8).Font.Bold = True
                            
                            xl.Cells(NFil, 9).Value = NVentaD
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 10).Value = NIGVD
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 11).Value = NTotalD
                            xl.Cells(NFil, 11).Font.Bold = True
                            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
                        End If
                    ElseIf optTipo(1).Value Then
                        
                        If txtCod_Moneda.Text <> "" Then
                            xl.Cells(NFil, 7).Value = "Sub Total"
                            xl.Cells(NFil, 7).Font.Bold = True
                            
                            xl.Cells(NFil, 8).Value = NVenta
                            xl.Cells(NFil, 8).Font.Bold = True
                            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 9).Value = NIGV
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                            
                            xl.Cells(NFil, 10).Value = NTotal
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                        Else
                            xl.Cells(NFil, 7).Value = "Sub Total Soles"
                            xl.Cells(NFil, 7).Font.Bold = True
                            
                            xl.Cells(NFil, 8).Value = NVentaS
                            xl.Cells(NFil, 8).Font.Bold = True
                            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 9).Value = NIGVS
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                            
                            xl.Cells(NFil, 10).Value = NTotalS
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                            
                            NFil = NFil + 1
                            
                            xl.Cells(NFil, 7).Value = "Sub Total Dolares"
                            xl.Cells(NFil, 7).Font.Bold = True
                            
                            xl.Cells(NFil, 8).Value = NVentaD
                            xl.Cells(NFil, 8).Font.Bold = True
                            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 9).Value = NIGVD
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                            
                            xl.Cells(NFil, 10).Value = NTotalD
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                        End If
                    ElseIf optTipo(2).Value Then
                        
                        If txtCod_Moneda.Text <> "" Then
                            xl.Cells(NFil, 7).Value = "Sub Total"
                            xl.Cells(NFil, 7).Font.Bold = True
                            
                            xl.Cells(NFil, 8).Value = NVenta
                            xl.Cells(NFil, 8).Font.Bold = True
                            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 9).Value = NIGV
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                            
                            xl.Cells(NFil, 10).Value = NTotal
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                        Else
                            xl.Cells(NFil, 7).Value = "Sub Total Soles"
                            xl.Cells(NFil, 7).Font.Bold = True
                            
                            xl.Cells(NFil, 8).Value = NVentaS
                            xl.Cells(NFil, 8).Font.Bold = True
                            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 9).Value = NIGVS
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                            
                            xl.Cells(NFil, 10).Value = NTotalS
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                            
                            NFil = NFil + 1
                            
                            xl.Cells(NFil, 7).Value = "Sub Total Dolares"
                            xl.Cells(NFil, 7).Font.Bold = True
                            
                            xl.Cells(NFil, 8).Value = NVentaD
                            xl.Cells(NFil, 8).Font.Bold = True
                            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
                    
                            xl.Cells(NFil, 9).Value = NIGVD
                            xl.Cells(NFil, 9).Font.Bold = True
                            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
                            
                            xl.Cells(NFil, 10).Value = NTotalD
                            xl.Cells(NFil, 10).Font.Bold = True
                            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
                            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
                        End If
                        
                    End If
                    
                    NFil = NFil + 1
                    
                End If
                        
            End If
            
        Loop
        
    End If
    
    RsC.Close: Set RsC = Nothing
    
    If optTipo(0).Value Then
                        
        If txtCod_Moneda.Text <> "" Then
            xl.Cells(NFil, 8).Value = "Sub Total"
            xl.Cells(NFil, 8).Font.Bold = True
            
            xl.Cells(NFil, 9).Value = NVenta
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 10).Value = NIGV
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 11).Value = NTotal
            xl.Cells(NFil, 11).Font.Bold = True
            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
        Else
            xl.Cells(NFil, 8).Value = "Sub Total Soles"
            xl.Cells(NFil, 8).Font.Bold = True
            
            xl.Cells(NFil, 9).Value = NVentaS
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 10).Value = NIGVS
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 11).Value = NTotalS
            xl.Cells(NFil, 11).Font.Bold = True
            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
            
            NFil = NFil + 1
            
            xl.Cells(NFil, 8).Value = "Sub Total Dolares"
            xl.Cells(NFil, 8).Font.Bold = True
            
            xl.Cells(NFil, 9).Value = NVentaD
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 10).Value = NIGVD
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 11).Value = NTotalD
            xl.Cells(NFil, 11).Font.Bold = True
            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
        End If
    ElseIf optTipo(1).Value Then
        
        If txtCod_Moneda.Text <> "" Then
            xl.Cells(NFil, 7).Value = "Sub Total"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVenta
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 9).Value = NIGV
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotal
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        Else
            xl.Cells(NFil, 7).Value = "Sub Total Soles"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaS
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 9).Value = NIGVS
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalS
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            NFil = NFil + 1
            
            xl.Cells(NFil, 7).Value = "Sub Total Dolares"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaD
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 9).Value = NIGVD
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalD
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        End If
    ElseIf optTipo(2).Value Then
        
        If txtCod_Moneda.Text <> "" Then
            xl.Cells(NFil, 7).Value = "Sub Total"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVenta
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 9).Value = NIGV
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotal
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        Else
            xl.Cells(NFil, 7).Value = "Sub Total Soles"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaS
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 9).Value = NIGVS
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalS
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            NFil = NFil + 1
            
            xl.Cells(NFil, 7).Value = "Sub Total Dolares"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaD
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
    
            xl.Cells(NFil, 9).Value = NIGVD
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalD
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        End If
        
    End If
    
    NFil = NFil + 1
    
    If optTipo(0).Value Then
    
        If txtCod_Moneda.Text <> "" Then
            xl.Cells(NFil, 8).Value = "Total General"
            xl.Cells(NFil, 8).Font.Bold = True
            
            xl.Cells(NFil, 9).Value = NVentaTot
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NIgvTot
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 11).Value = NTotalTot
            xl.Cells(NFil, 11).Font.Bold = True
            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
        Else
            xl.Cells(NFil, 8).Value = "Total General Soles"
            xl.Cells(NFil, 8).Font.Bold = True
            
            xl.Cells(NFil, 9).Value = NVentaTotS
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NIGVTotS
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 11).Value = NTotalTotS
            xl.Cells(NFil, 11).Font.Bold = True
            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
            
            NFil = NFil + 1
            
            xl.Cells(NFil, 8).Value = "Total General Dolares"
            xl.Cells(NFil, 8).Font.Bold = True
            
            xl.Cells(NFil, 9).Value = NVentaTotD
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NIGVTotD
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 11).Value = NTotalTotD
            xl.Cells(NFil, 11).Font.Bold = True
            xl.Cells(NFil, 11).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 11).HorizontalAlignment = xlRight
        End If
        
    ElseIf optTipo(1).Value Then
        
        If txtCod_Moneda.Text <> "" Then
            xl.Cells(NFil, 7).Value = "Total General"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaTot
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 9).Value = NIgvTot
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalTot
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        Else
            xl.Cells(NFil, 7).Value = "Total General Soles"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaTotS
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 9).Value = NIGVTotS
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalTotS
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            NFil = NFil + 1
            
            xl.Cells(NFil, 7).Value = "Total General Dolares"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaTotD
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 9).Value = NIGVTotD
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalTotD
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        End If
        
    ElseIf optTipo(2).Value Then
        
        If txtCod_Moneda.Text <> "" Then
            xl.Cells(NFil, 7).Value = "Total General"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaTot
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 9).Value = NIgvTot
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalTot
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        Else
            xl.Cells(NFil, 7).Value = "Total General Soles"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaTotS
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 9).Value = NIGVTotS
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalTotS
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
            
            NFil = NFil + 1
            
            xl.Cells(NFil, 7).Value = "Total General Dolares"
            xl.Cells(NFil, 7).Font.Bold = True
            
            xl.Cells(NFil, 8).Value = NVentaTotD
            xl.Cells(NFil, 8).Font.Bold = True
            xl.Cells(NFil, 8).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 8).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 9).Value = NIGVTotD
            xl.Cells(NFil, 9).Font.Bold = True
            xl.Cells(NFil, 9).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 9).HorizontalAlignment = xlRight
            
            xl.Cells(NFil, 10).Value = NTotalTotD
            xl.Cells(NFil, 10).Font.Bold = True
            xl.Cells(NFil, 10).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            xl.Cells(NFil, 10).HorizontalAlignment = xlRight
        End If
        
    End If
    
    xl.Visible = True
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
        
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    txtGls_Moneda.Text = "MONEDA ORIGINAL"
        
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
        
End Sub

Private Sub txtCod_Cliente_Change()
 
    If txtCod_Cliente.Text <> "" Then
        txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    Else
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If
    
End Sub

Private Sub txtCod_Moneda_Change()
    
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    Else
        txtGls_Moneda.Text = "MONEDA ORIGINAL"
    End If

End Sub

Private Sub txtCod_Producto_Change()
    
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If

End Sub
