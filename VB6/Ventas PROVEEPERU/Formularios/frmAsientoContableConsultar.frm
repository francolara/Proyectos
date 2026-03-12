VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAsientoContableConsultar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Asientos Contables"
   ClientHeight    =   8145
   ClientLeft      =   1650
   ClientTop       =   1530
   ClientWidth     =   13200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   13200
   Begin VB.Frame Frame1 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   645
      Width           =   13155
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7110
         Left            =   45
         OleObjectBlob   =   "frmAsientoContableConsultar.frx":0000
         TabIndex        =   1
         Top             =   165
         Width           =   13005
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   315
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
            Picture         =   "frmAsientoContableConsultar.frx":4F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":52CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":571E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":5AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":5E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":61EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":6586
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":6920
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":6CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":7054
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":73EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":80B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsientoContableConsultar.frx":844A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAsientoContableConsultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_empresa     As New ADODB.Connection
Dim cconex_empresa  As String

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

    ConfGrid GLista, False, False, False, False
    
    Me.Height = 8715
    Me.Width = 13440
    
    llenargrid StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    
    Exit Sub
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
    
    Select Case Button.Index
        Case 1
            imprimir StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2
            Unload Me
    End Select
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    
End Sub

Private Sub llenargrid(StrMsgError As String)
Dim rsLista As New ADODB.Recordset
Dim cselect As String
On Error GoTo Err
    
    If rsAsientosContables.State = 1 Then
        If rsAsientosContables.RecordCount = 0 Then
            StrMsgError = "No se han generado los asientos contables. Verifique"
            GoTo Err
        Else
            GLista.DefaultFields = False
    
            GLista.Dataset.ADODataset.CursorLocation = clUseClient
            GLista.Dataset.Active = False
               
            Set GLista.DataSource = rsAsientosContables
        
            GLista.Dataset.Active = True
            GLista.KeyField = "IdComprobante"
            If StrMsgError <> "" Then GoTo Err
        End If
    Else
        StrMsgError = "No se han generado los asientos contables. Verifique"
        GoTo Err
    End If
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub imprimir(StrMsgError As String)
On Error GoTo Err
    
    GLista.m.ExportToXLS App.Path & "\Temporales\AsientosContables.xls"
    ShellEx App.Path & "\Temporales\AsientosContables.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
            
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

