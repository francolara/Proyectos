VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form frmConfImpRecibos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configura Impresion Recibos"
   ClientHeight    =   6030
   ClientLeft      =   4050
   ClientTop       =   1305
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   5925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9390
      Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
         Height          =   5490
         Left            =   150
         OleObjectBlob   =   "frmConfImpRecibos.frx":0000
         TabIndex        =   1
         Top             =   300
         Width           =   9060
      End
   End
End
Attribute VB_Name = "frmConfImpRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0
        
    ConfGrid gCabecera, True, False, False, False
    listar
    
End Sub

Private Sub listar()

    csql = "Select Identificador,GlsObs,indImprime,impX,impY,impLongitud From objimprecibos Where idEmpresa = '" & glsEmpresa & "' AND trim(GlsCampo) <> ''"
    With gCabecera
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
End Sub
