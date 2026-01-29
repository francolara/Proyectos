VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmRankingVentasConsulta 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ranking de Ventas"
   ClientHeight    =   8760
   ClientLeft      =   1575
   ClientTop       =   1230
   ClientWidth     =   12930
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
   ScaleHeight     =   8760
   ScaleWidth      =   12930
   Begin DXDBGRIDLibCtl.dxDBGrid g 
      Height          =   7485
      Left            =   45
      OleObjectBlob   =   "frmRankingVentasConsulta.frx":0000
      TabIndex        =   0
      Top             =   1125
      Width           =   12795
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   11385
      Top             =   720
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
            Picture         =   "frmRankingVentasConsulta.frx":3382
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":371C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":3B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":3F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":42A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":463C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":49D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":4D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":510A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":54A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":583E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":6500
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":689A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":6CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":7086
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRankingVentasConsulta.frx":7A98
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RANKING DE VENTAS"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   720
      Width           =   12840
   End
End
Attribute VB_Name = "frmRankingVentasConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrFechaIni                         As String
Dim StrFechaFin                         As String
Dim CMon                                As String
Dim cSuc                                As String
Dim strPoficial                         As String
Dim CGlsReporte                         As String
Dim GlsForm                             As String
Dim cGrupo                              As String
Dim cNiveles                            As String
Dim X                                   As Integer
Dim COrden                              As String
Dim CIndMuestras                        As String

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
Dim CSqlOrden                           As String
Dim CSqlMuestras                        As String

    Me.left = 0
    Me.top = 0
    
    ConfGrid G, False, False, False, False
    If StrMsgError <> "" Then GoTo Err
    
    If COrden = "D" Then
        
        CSqlOrden = "GlsCliente"
    
    Else
    
        CSqlOrden = "TotalValorVenta Desc"
    
    End If
    
    If CIndMuestras = "1" Then
        
        CSqlMuestras = "And D.IndVtaGratuita = 1 "
    
    Else
        
        CSqlMuestras = "And D.IndVtaGratuita <> 1 "
       
    End If
    
    With G
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        CGlsReporte = "rptRankingVentasPorCliente.rpt"
        .Dataset.ADODataset.CommandText = "Call spu_RankingdeVentasPorCliente ('" & glsEmpresa & "','" & cSuc & "','" & CMon & "','" & StrFechaIni & "','" & StrFechaFin & "','','" & IIf(strPoficial, strPoficial, "1") & "','" & CSqlOrden & "','" & CSqlMuestras & "')"
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
    Exit Sub
 

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarDatos(ByVal pini As String, ByVal pfin As String, ByVal pmoneda As String, ByVal PTipo As String, psucursal As String, poficial As String, ByRef StrMsgError As String, POrden As String, PIndMuestras As String)
On Error GoTo Err

    StrFechaIni = Format(pini, "yyyy-mm-dd")
    StrFechaFin = Format(pfin, "yyyy-mm-dd")
    CMon = IIf(Trim(pmoneda) = "", "PEN", pmoneda)
    cSuc = psucursal
    strPoficial = poficial
    COrden = POrden
    CIndMuestras = PIndMuestras
    
    GlsForm = "Reporte de Ranking de Ventas - " & IIf(PTipo, " Por Cliente", " Por Producto")
    For X = 1 To glsNumNiveles
        cNiveles = cNiveles & "idNivel" & Format(X, "00") & ", GlsNivel" & Format(X, "00") & ","
    Next X
    
    Me.Show
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim Fini As String
Dim Ffin As String
Dim CodMoneda As String
Dim CGlsReporte     As String
Dim GlsForm         As String
Dim cGrupo          As String
Dim StrMsgError As String
Dim CSqlOrden                           As String
Dim CSqlMuestras                        As String
    
    If COrden = "D" Then
        
        CSqlOrden = "GlsCliente"
    
    Else
    
        CSqlOrden = "TotalValorVenta Desc"
    
    End If
    
    If CIndMuestras = "1" Then
        
        CSqlMuestras = "And D.IndVtaGratuita = 1 "
    
    Else
        
        CSqlMuestras = "And D.IndVtaGratuita <> 1 "
       
    End If
    
    Select Case Button.Index
        Case 1:
            CGlsReporte = "rptRankingVentasPorCliente.rpt"
            mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta|parCliente|ParOficial|ParOrden|parMuestras", glsEmpresa & "|" & cSuc & "|" & CMon & "|" & StrFechaIni & "|" & StrFechaFin & "|" & "" & "|" & IIf(strPoficial, strPoficial, "1") & "|" & CSqlOrden & "|" & CSqlMuestras, GlsForm, StrMsgError
        Case 2:
            G.m.ExportToXLS App.Path & "\Temporales\RankingVentas.xls"
            ShellEx App.Path & "\Temporales\RankingVentas.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 3:
            Unload Me
    End Select
    
    If StrMsgError <> "" Then GoTo Err
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
