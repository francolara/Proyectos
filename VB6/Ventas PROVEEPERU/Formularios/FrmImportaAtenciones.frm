VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmImportaAtenciones 
   Caption         =   "Ayuda de Atenciones"
   ClientHeight    =   8205
   ClientLeft      =   2220
   ClientTop       =   2295
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10995
   Begin VB.Frame FraLista 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10950
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   10740
         Begin CATControls.CATTextBox Txt_TextoBuscar 
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Top             =   210
            Width           =   4800
            _ExtentX        =   8467
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
            Container       =   "FrmImportaAtenciones.frx":0000
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin MSComCtl2.DTPicker DtpFechaD 
            Height          =   285
            Left            =   6750
            TabIndex        =   5
            Top             =   225
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
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
            Format          =   104792065
            CurrentDate     =   41713
         End
         Begin MSComCtl2.DTPicker DtpFechaH 
            Height          =   285
            Left            =   9360
            TabIndex        =   6
            Top             =   225
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
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
            Format          =   104792065
            CurrentDate     =   41713
         End
         Begin VB.Label Label57 
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
            Height          =   240
            Left            =   6075
            TabIndex        =   8
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label58 
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
            Height          =   240
            Left            =   8685
            TabIndex        =   7
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label56 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   210
            Width           =   765
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7080
         Left            =   135
         OleObjectBlob   =   "FrmImportaAtenciones.frx":001C
         TabIndex        =   4
         Top             =   945
         Width           =   10740
      End
   End
End
Attribute VB_Name = "FrmImportaAtenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CIdAtencion                 As String
Dim CIdCliente                  As String

Private Sub DtpFechaD_Change()
On Error GoTo Err
Dim StrMsgError                 As String

    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub DtpFechaH_Change()
On Error GoTo Err
Dim StrMsgError                 As String

    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                 As String

    ConfGrid GLista, False, False, False, False
    
    DtpFechaH.Value = getFechaSistema
    DtpFechaD.Value = "01/" & Format(Month(DtpFechaH.Value), "00") & "/" & Year(DtpFechaH)
    
    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    CIdAtencion = ""
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Lista(StrMsgError As String)
On Error GoTo Err
Dim strCond                     As String

    strCond = ""
    
    If Trim(txt_TextoBuscar.Text) <> "" Then
    
        strCond = "And (A.IdAtencion Like '%" & Trim(txt_TextoBuscar.Text) & "%' Or B.GlsPersona Like '%" & Trim(txt_TextoBuscar.Text) & "%' Or A.GlsNombres Like '%" & Trim(txt_TextoBuscar.Text) & "%' Or A.GlsApellidos Like '%" & Trim(txt_TextoBuscar.Text) & "%' Or A.IdAutorizacion Like '%" & Trim(txt_TextoBuscar.Text) & "%') "
        
    End If
    
    If Trim(CIdCliente) <> "" Then
    
        strCond = strCond & "And A.IdCliente = '" & CIdCliente & "' "
        
    End If
    
    csql = "Select A.IdAtencion,Date_Format(A.Fecha,'%d/%m/%Y') Fecha,B.GlsPersona,ConCat(A.GlsNombres,' ',A.GlsApellidos) Paciente,A.IdAutorizacion " & _
           "From RegistroAtenciones A " & _
           "Inner Join Personas B On A.IdCliente = B.IdPersona " & _
           "Left Join(" & _
               "Select Z.IdEmpresa,V.NumDocReferencia " & _
               "From DocVentas Z " & _
               "Left Join DocReferencia V " & _
                   "On Z.IdEmpresa = V.IdEmpresa And Z.IdSucursal = V.IdSucursal And Z.IdDocumento = V.TipoDocOrigen And Z.IdSerie = V.SerieDocOrigen " & _
                   "And Z.IdDocVentas = V.NumDocOrigen " & _
               "Where Z.IdEmpresa = '01' And Z.EstDocVentas In('GEN','IMP') " & _
           ") D " & _
               "On A.IdEmpresa = D.IdEmpresa And A.IdAtencion = D.NumDocReferencia " & _
           "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' " & _
           "And A.Fecha BetWeen '" & Format(DtpFechaD.Value, "yyyy-mm-dd") & "' And '" & Format(DtpFechaH.Value, "yyyy-mm-dd") & "' And A.IndEstado = 'C' " & _
           "And D.NumDocReferencia Is Null " & _
           "Group By A.IdAtencion " & _
           "Order By A.IdAtencion"

    'csql = "Select A.IdAtencion,B.GlsPersona,ConCat(A.GlsNombres,' ',A.GlsApellidos) Paciente,A.IdAutorizacion,A.Fecha " & _
           "From RegistroAtenciones A " & _
           "Inner Join Personas B " & _
               "On A.IdCliente = B.IdPersona " & _
           "Left Join DocReferencia C " & _
               "On A.IdEmpresa = C.IdEmpresa And A.IdSucursal = C.IdSucursal And '80' = C.TipoDocReferencia And '999' = C.SerieDocReferencia " & _
               "And A.IdAtencion = C.NumDocReferencia " & _
           "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' " & _
           "And A.Fecha BetWeen '" & Format(DtpFechaD.Value, "yyyy-mm-dd") & "' And '" & Format(DtpFechaH.Value, "yyyy-mm-dd") & "' And A.IndEstado = 'C' " & _
           "And C.NumDocReferencia Is Null " & strCond & _
           "Group By A.IdAtencion " & _
           "Order By A.IdAtencion"
           
    With GLista
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "IdAtencion"
        
    End With
    
    Me.Refresh
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub MostrarForm(StrMsgError As String, PIdAtencion As String, PIdCliente As String)
On Error GoTo Err
    
    CIdCliente = PIdCliente
    
    Me.Show 1
    
    PIdAtencion = CIdAtencion
    
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError                     As String
    
    CIdAtencion = Trim("" & GLista.Columns.ColumnByFieldName("IdAtencion").Value)
    Me.Hide

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError                 As String

    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
