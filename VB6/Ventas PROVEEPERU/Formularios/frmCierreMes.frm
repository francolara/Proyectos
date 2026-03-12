VERSION 5.00
Begin VB.Form frmCierreMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Mes"
   ClientHeight    =   2805
   ClientLeft      =   6150
   ClientTop       =   2070
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5700
   Begin VB.ComboBox cbxAno 
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
      ItemData        =   "frmCierreMes.frx":0000
      Left            =   1575
      List            =   "frmCierreMes.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2340
   End
   Begin VB.ComboBox cbxMes 
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
      ItemData        =   "frmCierreMes.frx":0050
      Left            =   1575
      List            =   "frmCierreMes.frx":0078
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   750
      Width           =   2340
   End
   Begin VB.Frame fraBotones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   75
      TabIndex        =   4
      Top             =   1950
      Width           =   5565
      Begin VB.CommandButton cmbOperar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1455
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmbCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.Label lbl_Estado 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   600
      TabIndex        =   5
      Top             =   1275
      Width           =   4440
   End
End
Attribute VB_Name = "frmCierreMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbxAno_Click()
Dim strAno As String
Dim strMes As String

    strAno = Format(cbxAno.Text, "0000")
    strMes = Format(CbxMes.ListIndex + 1, "00")
    ubicaDatos strAno, strMes

End Sub

Private Sub cbxMes_Click()
Dim strAno As String
Dim strMes As String

    strAno = Format(cbxAno.Text, "0000")
    strMes = Format(CbxMes.ListIndex + 1, "00")
    ubicaDatos strAno, strMes

End Sub

Private Sub cmbCancelar_Click()
    
    Unload Me

End Sub

Private Sub cmbOperar_Click()
On Error GoTo Err
Dim StrMsgError As String

    If MsgBox("¿Seguro de realizar el proceso?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        abrirCerrar StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim fecha As Date
Dim strAno As String
Dim strMes As String
Dim i As Integer
    
    Me.top = 0
    Me.left = 0
    
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
    CbxMes.ListIndex = Val(strMes) - 1
    ubicaDatos strAno, strMes

End Sub

Private Sub ubicaDatos(ByVal strVarAno As String, ByVal strVarMes As String)
Dim strEstado As String

    strEstado = traerCampo("cierresmes", "estCierre", "idAno", strVarAno, True, "idMes = '" & strVarMes & "' And IdSistema = '21001'")
    If strEstado = "" Or strEstado = "A" Then
        lbl_Estado.Caption = "ABIERTO"
        cmbOperar.Caption = "Cerrar"
    Else
        lbl_Estado.Caption = "CERRADO"
        cmbOperar.Caption = "Abrir"
    End If

End Sub

Private Sub abrirCerrar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strAno As String
Dim strMes As String
Dim strEstado As String
Dim strEstadoActual As String
Dim strMsg As String

    strAno = cbxAno.Text
    strMes = Format(CbxMes.ListIndex + 1, "00")
    
    strEstado = "C"
    strMsg = "Cerro"
    
    If cmbOperar.Caption = "Abrir" Then
        
        If traerCampo("CierresMes", "EstCierre", "IdAno", strAno, True, "IdMes = '" & strMes & "' And IdSistema = '21008'") = "C" Then
            
            StrMsgError = "El Mes se encuentra cerrado en Contabilidad, no puede abrir el mes.": GoTo Err
        
        End If
        
    End If
    
    If cmbOperar.Caption = "Abrir" Then
        strEstado = "A"
        strMsg = "Abrio"
    End If
    
    strEstadoActual = traerCampo("cierresmes", "estCierre", "idAno", strAno, True, "idMes = '" & strMes & "' And IdSistema = '21001'")
    If strEstadoActual = "" Then
        csql = "INSERT INTO cierresmes (idEmpresa,FecCierre,idMes,idAno,estCierre,IdSistema) VALUES ('" & glsEmpresa & "',GETDATE(),'" & strMes & "','" & strAno & "','" & strEstado & "','21001')"
    Else
        csql = "UPDATE cierresmes SET FecCierre = GETDATE(),estCierre = '" & strEstado & "' WHERE idEmpresa = '" & glsEmpresa & "' AND idMes = '" & strMes & "' AND idAno = '" & strAno & "' And IdSistema = '21001'"
    End If
    Cn.Execute csql
    
    ubicaDatos strAno, strMes
    MsgBox "Se " & strMsg & " satisfactoriamente", vbInformation, App.Title
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
