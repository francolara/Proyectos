VERSION 5.00
Begin VB.Form frmAsientoContableTransferir 
   Caption         =   "Transferir Asientos Contables"
   ClientHeight    =   2130
   ClientLeft      =   5850
   ClientTop       =   3255
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5280
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
      Height          =   435
      Left            =   2655
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
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
      Height          =   435
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   90
      TabIndex        =   0
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
         ItemData        =   "frmAsientoContableTransferir.frx":0000
         Left            =   1710
         List            =   "frmAsientoContableTransferir.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   2
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
         ItemData        =   "frmAsientoContableTransferir.frx":0050
         Left            =   1710
         List            =   "frmAsientoContableTransferir.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   675
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
         Left            =   1035
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
         Left            =   1035
         TabIndex        =   4
         Top             =   720
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
         Left            =   1035
         TabIndex        =   3
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmAsientoContableTransferir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strcnConta      As String
Option Explicit
Dim strAno          As String
Dim strMes          As String
Dim docofi          As String
Dim IndGeneraInt                            As Boolean

Public Sub Genera_Internamente(StrMsgError As String, PPeriodo As String, PIndGeneraInt As Boolean)
On Error GoTo Err
Dim i                           As Long

    Load Me
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = left(PPeriodo, 4) Then Exit For
    Next
    
    cbxMes.ListIndex = Val(right(PPeriodo, 2)) - 1
    
    IndGeneraInt = True
    
    cmbOperar_Click
    
    PIndGeneraInt = IndGeneraInt
    
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub cmbCancelar_Click()

    Unload Me

End Sub

Private Sub cmbOperar_Click()

    TRANSFIERE

End Sub

Private Sub TRANSFIERE()
On Error GoTo Err
Dim strAno                  As String
Dim strMes                  As String
Dim cnn_empresa             As New ADODB.Connection
Dim rsasiento               As New ADODB.Recordset
Dim cconex_empresa          As String
Dim CnConta                 As New ADODB.Connection
Dim cinsert                 As String
Dim cCompro                 As String

Dim strcnConta2             As String
Dim CnConta2                As New ADODB.Connection

Dim strIdDocumento          As String
Dim strIdSerie              As String
Dim strIdDocventas          As String
Dim strIdComprobante        As String

strIdDocumento = ""
strIdSerie = ""
strIdDocventas = ""
strIdComprobante = ""

  docofi = traerCampo("Parametros", "ValParametro", "GlsParametro", "VISUALIZA_FILTRO_DOCUMENTO", True)

 If right(CmbOpciones.Text, 2) = "02" Then
    strcnConta2 = "dsn=dnsContabilidad2"
    
    CnConta2.CursorLocation = adUseClient
    CnConta2.Open strcnConta2
 End If


    strcnConta = "dsn=dnsContabilidad"
    
    CnConta.CursorLocation = adUseClient
    CnConta.Open strcnConta
        
        
    Me.MousePointer = 11

    strAno = cbxAno.Text
    strMes = Format(cbxMes.ListIndex + 1, "00")

    'If cnn_empresa.State = adStateOpen Then cnn_empresa.Close
    'cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Temporales\DB_CONTA.MDB" & ";Persist Security Info=False"
    'cnn_empresa.Open cconex_empresa
    
    'If rsasiento.State = adStateOpen Then rsasiento.Close
    'rsasiento.Open "SELECT * FROM CONTABLE ORDER BY f3compro,F3ELEMEN", cnn_empresa, adOpenDynamic, adLockOptimistic
    '----------------------------------------------------------------------------------------------------------
    
    'If Not rsAsientosContables.EOF Then
    If rsAsientosContables.State = 1 Then
        If rsAsientosContables.RecordCount = 0 Then
            MsgBox "No se han generado los asientos contables. Verifique.", vbInformation, App.Title
            GoTo Err
        Else
            rsAsientosContables.MoveFirst
            If rsAsientosContables.Fields("idPeriodo") = strAno & strMes Then
                If Trim("" & rsAsientosContables.Fields("TipoContable")) = right(CmbOpciones.Text, 2) Then
                    Do While Not rsAsientosContables.EOF
                        cCompro = rsAsientosContables.Fields("idComprobante")
                        
                        strIdDocumento = UCase(rsAsientosContables.Fields("idTipoDoc") & "")
                        strIdSerie = Trim(rsAsientosContables.Fields("SerieDoc") & "") 'Format(rsAsientosContables.Fields("SerieDoc") & "", "0000")
                        strIdDocventas = Format(rsAsientosContables.Fields("NumCheque") & "", "00000000")
                        strIdComprobante = rsAsientosContables.Fields("idComprobante") & ""
                        
                        
                        cinsert = "Insert Into AsientoContable " & _
                                  "(IdPeriodo,IdComprobante,IdOrigen,FecEmision,ValTc,IdEmpresa,IdMoneda,Estado,glsGlosa,FecRegistro,IdUsuarioReg,GlsPC,GlsPCUsuario) " & _
                                  "Values " & _
                                  "('" & rsAsientosContables.Fields("idPeriodo") & "','" & rsAsientosContables.Fields("idComprobante") & "','" & rsAsientosContables.Fields("IDOrigen") & _
                                  "','" & Format(rsAsientosContables.Fields("FecCompro"), "yyyy-mm-dd") & "'," & _
                                  rsAsientosContables.Fields("ValorTipoCambio") & ",'" & glsEmpresa & "','" & rsAsientosContables.Fields("idMoneda") & "','1'," & _
                                  "'" & rsAsientosContables.Fields("Glosa") & "',SysDate(),'" & glsUser & "','" & fpComputerName & "','" & fpUsuarioActual & "')"
                                  
                                If right(CmbOpciones.Text, 2) = "01" Then
                                    CnConta.Execute (cinsert)
                                Else
                                    CnConta2.Execute (cinsert)
                                End If
                        
                        Do While cCompro = rsAsientosContables.Fields("idComprobante") & ""
                            cinsert = "Insert Into AsientoContableDetalle " & _
                                      "(idComprobante,idPeriodo,Valitem,idOrigen,glsDetalle," & _
                                      "idCuentaContable,NumCheque,NumReferencia,TotalImporteS,TotalImporteD,idMoneda,ValTC," & _
                                      "IdTipoDH,CtaAuxiliar,IdEmpresa,idcosto,idtipodoc,seriedoc,FecCompro) " & _
                                      "Values" & _
                                      "('" & rsAsientosContables.Fields("idComprobante") & "','" & rsAsientosContables.Fields("idPeriodo") & "'," & rsAsientosContables.Fields("ValItem") & ",'" & rsAsientosContables.Fields("IDOrigen") & "','" & _
                                      rsAsientosContables.Fields("GlsDetalle") & "','" & rsAsientosContables.Fields("idCtaContable") & "'," & rsAsientosContables.Fields("NumCheque") & ",'" & _
                                      rsAsientosContables.Fields("NumReferencia") & "'," & rsAsientosContables.Fields("TotalImporteS") & "," & rsAsientosContables.Fields("TotalImporteD") & ",'" & rsAsientosContables.Fields("idMoneda") & "'" & _
                                      "," & rsAsientosContables.Fields("ValorTipoCambio") & ",'" & rsAsientosContables.Fields("idTipoDH") & "','" & rsAsientosContables.Fields("CtaAuxiliar") & _
                                      "','" & glsEmpresa & "','" & rsAsientosContables.Fields("idCosto") & "','" & rsAsientosContables.Fields("idTipoDoc") & "','" & rsAsientosContables.Fields("SerieDoc") & "','" & Format(rsAsientosContables.Fields("FecCompro"), "yyyy-mm-dd") & "')"
        
                                If right(CmbOpciones.Text, 2) = "01" Then
                                    CnConta.Execute (cinsert)
                                Else
                                    CnConta2.Execute (cinsert)
                                End If
                            
                            rsAsientosContables.MoveNext
                            If rsAsientosContables.EOF Then Exit Do
                        Loop
                        
                        '''' ACTUALIZA VENTAS'''''
                        If docofi = "S" Then
                            Select Case strIdDocumento
                                Case "FAC": strIdDocumento = "01"
                                Case "BOL": strIdDocumento = "03"
                                Case "CRE": strIdDocumento = "07"
                                Case "DEB": strIdDocumento = "08"
                                Case "T/C": strIdDocumento = "12"
                                Case "NPD": strIdDocumento = "90"
                            End Select
                         Else
                            Select Case strIdDocumento
                                Case "FAC": strIdDocumento = "01"
                                Case "BOL": strIdDocumento = "03"
                                Case "CRE": strIdDocumento = "07"
                                Case "DEB": strIdDocumento = "08"
                                Case "T/C": strIdDocumento = "12"
                            End Select
                         End If
                        
                        'If "00045973" = Trim(strIdDocventas) Then MsgBox ""
                
                        If right(CmbOpciones.Text, 2) = "01" Then
                            csql = "update docventas set indTrasladoConta = 'S', idComprobante = '" & strIdComprobante & "' " & _
                                     "where iddocumento = '" & strIdDocumento & "' and idserie = '" & strIdSerie & "' " & _
                                     "and iddocventas = '" & strIdDocventas & "' and idEmpresa = '" & glsEmpresa & "'"
                        Else
                            csql = "update docventas set indTrasladoContaFin = 'S', idComprobante = '" & strIdComprobante & "' " & _
                                     "where iddocumento = '" & strIdDocumento & "' and idserie = '" & strIdSerie & "' " & _
                                     "and iddocventas = '" & strIdDocventas & "' and idEmpresa = '" & glsEmpresa & "'"
                        End If
                        
                        Cn.Execute (csql)
                        
                        strIdSerie = ""
                        strIdDocventas = ""
                        strIdComprobante = ""
                        
                        ''''''''''''''''''''''''''
                        
                    Loop
                    
                    'actualizarDocVentas
            
                    If Not IndGeneraInt Then MsgBox "Se realizó la Transferencia de los Asientos Contables.", vbInformation, App.Title
                    rsAsientosContables.Close: Set rsasiento = Nothing
                Else
                    If Not IndGeneraInt Then MsgBox "Los asientos generados no corresponden a la Contabilidad que se desea transferir. Verifique", vbInformation, App.Title
                End If
            Else
                If Not IndGeneraInt Then MsgBox "Los asientos generados no corresponden al mes que se desea transferir. Verifique", vbInformation, App.Title
            End If
        End If
    Else
        If Not IndGeneraInt Then MsgBox "No se han generado los asientos contables. Verifique.", vbInformation, App.Title
    End If
    
    Me.MousePointer = 1

    CnConta.Close
    Exit Sub
Err:
    IndGeneraInt = False
    MsgBox Err.Description, vbInformation, App.Title
    CnConta.Close
    Exit Sub
End Sub

Private Sub Form_Load()
Dim fecha As Date
Dim i As Integer

    Me.Width = 5400
    Me.Height = 2670
    
    IndGeneraInt = False
    
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

Private Sub actualizarDocVentas()
Dim cnn_f5pla      As New ADODB.Connection
Dim tbcf5pla       As New ADODB.Recordset
Dim cconex_f5pla   As String
Dim strIdDocumento As String
    
    docofi = traerCampo("Parametros", "ValParametro", "GlsParametro", "VISUALIZA_FILTRO_DOCUMENTO", True)
    
    'If cnn_f5pla.State = adStateOpen Then cnn_f5pla.Close
    'cconex_f5pla = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Temporales\DB_CONTA.MDB" & ";Persist Security Info=False"
    'cnn_f5pla.Open cconex_f5pla

    'csql = "SELECT DISTINCT F3DETALL, F3TIPDOC, F3CHEQUE, F3SERDOC FROM CONTABLE"
    'If tbcf5pla.State = adStateOpen Then tbcf5pla.Close
    'tbcf5pla.Open csql, cnn_f5pla, adOpenStatic, adLockOptimistic
    
    If rsAsientosContables.RecordCount <> 0 Then
        rsAsientosContables.MoveFirst
        Do While Not rsAsientosContables.EOF
            strIdDocumento = ""
            
            If docofi = "S" Then
                Select Case UCase(rsAsientosContables.Fields("idTipoDoc") & "")
                    Case "FAC": strIdDocumento = "01"
                    Case "BOL": strIdDocumento = "03"
                    Case "CRE": strIdDocumento = "07"
                    Case "DEB": strIdDocumento = "08"
                    Case "T/C": strIdDocumento = "12"
                    Case "NPD": strIdDocumento = "90"
                End Select
             Else
                Select Case UCase(rsAsientosContables.Fields("idTipoDoc") & "")
                    Case "FAC": strIdDocumento = "01"
                    Case "BOL": strIdDocumento = "03"
                    Case "CRE": strIdDocumento = "07"
                    Case "DEB": strIdDocumento = "08"
                    Case "T/C": strIdDocumento = "12"
                End Select
             End If
            
    
            If right(CmbOpciones.Text, 2) = "01" Then
                csql = "update docventas set indTrasladoConta = 'S' " & _
                         "where iddocumento = '" & strIdDocumento & "' and idserie = '" & Trim(rsAsientosContables.Fields("SerieDoc") & "") & "' " & _
                         "and iddocventas = '" & Format(rsAsientosContables.Fields("NumCheque") & "", "00000000") & "' and idEmpresa = '" & glsEmpresa & "'"
            Else
                csql = "update docventas set indTrasladoContaFin = 'S' " & _
                         "where iddocumento = '" & strIdDocumento & "' and idserie = '" & Trim(rsAsientosContables.Fields("SerieDoc") & "") & "' " & _
                         "and iddocventas = '" & Format(rsAsientosContables.Fields("NumCheque") & "", "00000000") & "' and idEmpresa = '" & glsEmpresa & "'"
            End If
            
            Cn.Execute (csql)
            
            rsAsientosContables.MoveNext
        Loop
    End If
    
End Sub
