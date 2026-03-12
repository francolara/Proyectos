VERSION 5.00
Begin VB.Form FrmGeneraPDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PDB"
   ClientHeight    =   2115
   ClientLeft      =   3840
   ClientTop       =   4005
   ClientWidth     =   5520
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
   ScaleHeight     =   2115
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdTC 
      Caption         =   "TC"
      Height          =   435
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Ventas"
      Height          =   435
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   5415
      Begin VB.ComboBox cbxAno 
         Height          =   330
         ItemData        =   "FrmGeneraPDB.frx":0000
         Left            =   2115
         List            =   "FrmGeneraPDB.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2340
      End
      Begin VB.ComboBox cbxMes 
         Height          =   330
         ItemData        =   "FrmGeneraPDB.frx":0050
         Left            =   2115
         List            =   "FrmGeneraPDB.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   210
         Left            =   1395
         TabIndex        =   6
         Top             =   450
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   210
         Left            =   1395
         TabIndex        =   5
         Top             =   855
         Width           =   300
      End
   End
End
Attribute VB_Name = "FrmGeneraPDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    GeneraArchivo StrMsgError, True
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub GeneraArchivo(StrMsgError As String, PIndVentas As Boolean)
On Error GoTo Err
Dim CSqlC                           As String
Dim RsC                             As New ADODB.Recordset
    
    If PIndVentas Then
    
        CSqlC = "Call Spu_GeneraVentas_PDB('" & glsEmpresa & "','" & cbxAno.Text & "','" & Format(cbxMes.ListIndex + 1, "00") & "')"
        RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
        If Not RsC.EOF Then
            
            GeneraTxt StrMsgError, RsC, "V" & traerCampo("Empresas", "Ruc", "IdEmpresa", glsEmpresa, False) & cbxAno.Text & Format(cbxMes.ListIndex + 1, "00") & ".TXT"
            
        End If
        
    Else
    
        CSqlC = "Select Date_Format(Fecha,'%d/%m/%Y'),TcCompra,TcVenta " & _
                "From TiposDeCambio " & _
                "Where Year(Fecha) = '" & cbxAno.Text & "' And Month(Fecha) = " & cbxMes.ListIndex + 1 & ""
        RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
        If Not RsC.EOF Then
            
            GeneraTxt StrMsgError, RsC, traerCampo("Empresas", "Ruc", "IdEmpresa", glsEmpresa, False) & ".TC"
            
        End If
        
    End If
    
    RsC.Close: Set RsC = Nothing
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub cmdsalir_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdTC_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    GeneraArchivo StrMsgError, False
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                     As String
Dim fecha                           As Date
Dim i                               As Integer
Dim strAno                          As String
Dim strMes                          As String

    fecha = Format(getFechaSistema, "dd/mm/yyyy")
    strAno = Format(Year(fecha), "0000")
    strMes = Format(Month(fecha), "00")
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = strAno Then Exit For
    Next
    
    cbxMes.ListIndex = Val(strMes) - 1
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
