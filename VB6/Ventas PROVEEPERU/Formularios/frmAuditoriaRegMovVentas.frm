VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmAuditoriaRegMovVentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auditoría - Registro de Movimientos"
   ClientHeight    =   3405
   ClientLeft      =   4770
   ClientTop       =   2370
   ClientWidth     =   7365
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
   ScaleHeight     =   3405
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   3735
      TabIndex        =   5
      Top             =   2745
      Width           =   1290
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2385
      TabIndex        =   4
      Top             =   2745
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   7260
      Begin VB.CommandButton cmbAyudaSucursal 
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
         Picture         =   "frmAuditoriaRegMovVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1500
         Width           =   390
      End
      Begin VB.CommandButton cmdusuarios 
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
         Picture         =   "frmAuditoriaRegMovVentas.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1935
         Width           =   390
      End
      Begin VB.TextBox txtnomusuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "TODOS LOS USUARIOS"
         Top             =   1935
         Width           =   4785
      End
      Begin VB.TextBox txtusuario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   940
         TabIndex        =   3
         Top             =   1935
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Caption         =   " Rango de Fechas "
         Height          =   960
         Left            =   180
         TabIndex        =   7
         Top             =   315
         Width           =   6900
         Begin MSComCtl2.DTPicker dtpdesde 
            Height          =   315
            Left            =   1650
            TabIndex        =   0
            Top             =   405
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   103940097
            CurrentDate     =   38955
         End
         Begin MSComCtl2.DTPicker dtphasta 
            Height          =   315
            Left            =   4455
            TabIndex        =   1
            Top             =   405
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   103940097
            CurrentDate     =   38955
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   945
            TabIndex        =   9
            Top             =   465
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   3780
            TabIndex        =   8
            Top             =   465
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   940
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1485
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
         Container       =   "frmAuditoriaRegMovVentas.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   1890
         TabIndex        =   14
         Top             =   1485
         Width           =   4785
         _ExtentX        =   8440
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
         Container       =   "frmAuditoriaRegMovVentas.frx":0730
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   15
         Top             =   1530
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   1980
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmAuditoriaRegMovVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaSucursal_Click()

    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
    
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarReporte "RptAuditoriaRegMovVentas.rpt", "ParEmpresa|ParSucursal|ParFecDesde|ParFecHasta|ParUsuario", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Format(dtpdesde.Value, "yyyy-mm-dd") & "|" & Format(dtphasta.Value, "yyyy-mm-dd") & "|" & txtusuario.Text, "Reporte de Auditoría Registro de Movimientos", StrMsgError
    If StrMsgError <> "" Then GoTo Err
            
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub cmdusuarios_Click()

    txtusuario_KeyDown 113, 0
    
End Sub

Private Sub dtpdesde_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        dtphasta.SetFocus
    End If

End Sub

Private Sub dtphasta_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If Format(dtphasta.TabIndex, "dd/mm/yyyy") >= Format(dtpdesde.TabIndex, "dd/mm/yyyy") Then
            txtusuario.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_Load()

    dtpdesde.Value = Format(Date, "dd/mm/yyyy")
    dtphasta.Value = Format(Date, "dd/mm/yyyy")
    txtCod_Sucursal.Text = glsSucursal
    
End Sub

Private Sub txtusuario_Change()

    If Len(txtusuario.Text) = 0 Then
        txtusuario.Text = ""
        txtnomusuario.Text = "TODOS LOS USUARIOS"
    Else
        txtnomusuario.Text = traerCampo("Personas", "GlsPersona", "IdPersona", txtusuario.Text, False)
    End If
    
End Sub

Private Sub txtusuario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERROR

    If KeyCode = 113 Then
        Me.MousePointer = 11
        ayuda_usuarios.Show 1
        
        If Len(Trim("" & wusuario)) > 0 Then
            txtusuario.Text = Trim("" & wusuario)
        End If
        Me.MousePointer = 1
        txtusuario_KeyPress 13
    End If
    
    Exit Sub
 
ERROR:
   MsgBox "Se ha producido el sgte. error : " & Err.Description, vbCritical, App.Title
   Exit Sub
End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Sucursal_Change()

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
End Sub

Private Sub txtCod_Sucursal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Sucursal.Text = ""
    End If

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub
