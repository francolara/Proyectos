VERSION 5.00
Begin VB.Form frmAsientoContableEliminar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar Asientos Contables"
   ClientHeight    =   2055
   ClientLeft      =   6540
   ClientTop       =   2865
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "&Salir"
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
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1140
   End
   Begin VB.CommandButton cmbOperar 
      Caption         =   "Aceptar"
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
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   5100
      Begin VB.ComboBox CmbOpciones 
         Appearance      =   0  'Flat
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1035
         Visible         =   0   'False
         Width           =   2355
      End
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
         ItemData        =   "frmAsientoContableEliminar.frx":0000
         Left            =   1710
         List            =   "frmAsientoContableEliminar.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
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
         ItemData        =   "frmAsientoContableEliminar.frx":0050
         Left            =   1710
         List            =   "frmAsientoContableEliminar.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   2340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   6
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   5
         Top             =   315
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmAsientoContableEliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strAno      As String
Dim strMes      As String
Dim strcnConta  As String

Private Sub cmbCancelar_Click()

    Unload Me

End Sub

Private Sub cmbOperar_Click()
On Error GoTo Err
Dim StrMsgError As String

    If Trim("" & traerCampo("cierresmes", "estcierre", "Idmes", Format(cbxMes.ListIndex + 1, "00"), True, " idano = '" & cbxAno.Text & "' and IdSistema = '21008' ")) = "C" Then
        StrMsgError = "Contabilidad Se Encuentra Cerrado."
        GoTo Err
    End If


    If MsgBox("Està seguro(a) de Eliminar los asientos contables ?", vbQuestion + vbYesNo, "Atención") = vbYes Then
        ELIMINAR_VENTAS
    End If
    
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ELIMINAR_VENTAS()
On Error GoTo Err
Dim cperiodo        As String, strAno As String, strMes As String
Dim cdelete         As String, cupdate As String, corigen As String
Dim CnConta         As New ADODB.Connection
Dim strcnConta2     As String
Dim CnConta2        As New ADODB.Connection
Dim CSqlC           As String


    If right(CmbOpciones.Text, 2) = "02" Then
        strcnConta2 = "dsn=dnsContabilidad2"
        CnConta2.CursorLocation = adUseClient
        CnConta2.Open strcnConta2
    End If

    strcnConta = "dsn=dnsContabilidad"
    CnConta.CursorLocation = adUseClient
    CnConta.Open strcnConta

    Me.MousePointer = 11

    cperiodo = cbxAno.Text & Format(cbxMes.ListIndex + 1, "00")
    strAno = cbxAno.Text
    strMes = Format(cbxMes.ListIndex + 1, "00")
    corigen = traerCampo("parametros", "valparametro", "glsparametro", "ORIGEN_CONTABLE", True)
    
    If MsgBox("Se procederá a Eliminar los asientos contables. Desea continuar ?", vbQuestion + vbYesNo, "Atención") = vbYes Then
        
        cdelete = "DELETE FROM ASIENTOCONTABLE " & _
                  "WHERE idempresa = '" & glsEmpresa & "' and idperiodo = '" & cperiodo & "' and idorigen = '" & corigen & "'"
        If right(CmbOpciones.Text, 2) = "01" Then
            CnConta.Execute (cdelete)
        Else
            CnConta2.Execute (cdelete)
        End If
        
        cdelete = "DELETE FROM ASIENTOCONTABLEDETALLE " & _
                  "WHERE idempresa = '" & glsEmpresa & "' and idperiodo = '" & cperiodo & "' and idorigen = '" & corigen & "'"
        If right(CmbOpciones.Text, 2) = "01" Then
            CnConta.Execute (cdelete)
        Else
            CnConta2.Execute (cdelete)
        End If
                  
        If right(CmbOpciones.Text, 2) = "01" Then
            cupdate = "Update docventas set IndTrasladoConta = 'N', idComprobante  ='' " & _
                      "Where year(FecEmision)=" & Val(strAno) & " And Month(FecEmision)=" & Val(strMes) & _
                      " And IdEmpresa='" & glsEmpresa & "'"
        Else
            cupdate = "Update docventas set indTrasladoContaFin = 'N', idComprobante  ='' " & _
                      "Where year(FecEmision)=" & Val(strAno) & " And Month(FecEmision)=" & Val(strMes) & _
                      " And IdEmpresa='" & glsEmpresa & "'"
        End If
        Cn.Execute (cupdate)
        
        CSqlC = "Insert Into EliminaTransferencias(IdEmpresa,IdSistema,Periodo,IdUsuario,GlsPC,GlsPCUsuario,FechaHora)Values(" & _
                "'" & glsEmpresa & "','" & StrcodSistema & "','" & cperiodo & "','" & glsUser & "','" & fpComputerName & "','" & fpUsuarioActual & "',SysDate())"
        
        If right(CmbOpciones.Text, 2) = "01" Then
            CnConta.Execute (CSqlC)
        Else
            CnConta2.Execute (CSqlC)
        End If
        
        MsgBox "Se eliminaron los asientos contables. Verifique.", vbInformation, App.Title
    End If
    CnConta.Close
    
    Me.MousePointer = 1

    Exit Sub
    
Err:
    MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim fecha As Date
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
    cbxMes.ListIndex = Val(strMes) - 1
    
    CmbOpciones.AddItem "Tributaria" & Space(150) & "01"
    CmbOpciones.AddItem "Financiera" & Space(150) & "02"
    CmbOpciones.ListIndex = 0
    
    If Trim("" & traerCampo("Parametros", "Valparametro", "Glsparametro", "VISUALIZA_FILTRO_DOCUMENTO", True)) = "S" Then
        Label3.Visible = True
        CmbOpciones.Visible = True
    Else
        Label3.Visible = False
        CmbOpciones.Visible = False
    End If
    
End Sub
