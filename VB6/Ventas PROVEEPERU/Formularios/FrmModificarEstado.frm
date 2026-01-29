VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmModificarEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberar Documento"
   ClientHeight    =   2325
   ClientLeft      =   3735
   ClientTop       =   2280
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7245
   Begin VB.CommandButton BtnSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3735
      TabIndex        =   4
      Top             =   1755
      Width           =   1140
   End
   Begin VB.CommandButton Btnaceptar 
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
      Height          =   420
      Left            =   2520
      TabIndex        =   3
      Top             =   1755
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   7080
      Begin VB.CommandButton cmbAyudaTipoDoc 
         Height          =   315
         Left            =   6535
         Picture         =   "FrmModificarEstado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         Width           =   390
      End
      Begin VB.TextBox txtserie 
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
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   675
         Width           =   1140
      End
      Begin VB.TextBox txtnumdoc 
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
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   1035
         Width           =   1140
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   315
         Width           =   1140
         _ExtentX        =   2011
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
         Container       =   "FrmModificarEstado.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   2790
         TabIndex        =   10
         Top             =   315
         Width           =   3725
         _ExtentX        =   6562
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
         Container       =   "FrmModificarEstado.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número"
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
         Left            =   300
         TabIndex        =   9
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Left            =   300
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
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
         Left            =   300
         TabIndex        =   7
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmModificarEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAceptar_Click()
On Error GoTo Err
Dim StrMsgError As String

    If Len(txtCod_Documento.Text) = 0 Or Len(txtCod_Documento.Text) > 2 Then
        MsgBox "Tipo de Documento Incorrecto,Verifque", vbInformation, App.Title
        txtCod_Documento.SetFocus
        Exit Sub
    End If

    If Len(txtserie.Text) = 0 Or Len(txtserie.Text) > 4 Then
        MsgBox "Numero de Serie Incorrecto,Verifque", vbInformation, App.Title
        txtserie.SetFocus
        Exit Sub
    End If
    
    If Len(txtnumdoc.Text) = 0 Or Len(txtnumdoc.Text) > 8 Then
        MsgBox "Numero de Documento Incorrecto,Verifque", vbInformation, App.Title
        txtnumdoc.SetFocus
        Exit Sub
    End If
 
    
    validar "A", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub BtnSalir_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub actualizar(StrMsgError As String)
On Error GoTo Err
Dim rsg             As New ADODB.Recordset
Dim estado          As String
Dim idFormaPago     As String
Dim cabrev          As String
Dim cdocumento      As String
Dim Cta_Dcto        As String
Dim Pagos           As Integer
Dim srtValesCab     As String
Dim srtEstDoc       As String
 
    estado = "GEN"
    If MsgBox("¿Seguro de liberar el documento?", vbInformation + vbYesNo, App.Title) = vbYes Then
'        If txtCod_Documento.Text = "86" Then
'            csql = "select dv.idDocumento, dv.idSerie, dv.idDocVentas, dv.estDocVentas " & _
'                     "from docreferencia dr, docventas dv " & _
'                     "where dr.TipoDocReferencia = dv.idDocumento " & _
'                     "and dr.serieDocReferencia = dv.idSerie " & _
'                     "and dr.numDocReferencia = dv.idDocVentas " & _
'                     "and dr.idEmpresa = dv.idEmpresa " & _
'                     "and dr.TipoDocOrigen = '86' " & _
'                     "and dr.serieDocOrigen = '" & txtserie.Text & "' " & _
'                     "and dr.numDocOrigen = '" & txtnumdoc.Text & "' " & _
'                     "and dr.TipoDocReferencia = '01' " & _
'                     "and dr.idEmpresa = '" & glsEmpresa & "' "
'            If rsg.State = 1 Then rsg.Close
'            rsg.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'
'            If rsg.RecordCount <> 0 Then
'                rsg.MoveFirst
'                '--- Verificamos el estado de la factura
'                If rsg.Fields("estDocVentas") <> "ANU" Then
'                    StrMsgError = "La factura generada no ha sido anulada, no se puede habilitar la Guia"
'                    GoTo Err
'                End If
'
'                '--- Verificamos si la factura tiene pagos
'                idFormaPago = traerCampo("movcajasdet", "idFormadePago", "iddocventas", rsg.Fields("idDocVentas"), True, "idSerie = '" & rsg.Fields("idSerie") & "' and idDocumento = '01' and idTipoMovCaja = '99990002'")
'                cabrev = "Fac"
'                cdocumento = cabrev & serieFactura & "/" & numFactura
'                Cta_Dcto = traerCampo("Cta_Dcto", "idCta_Dcto", "Nro_Comp", cdocumento, True)
'                Pagos = Val(traerCampo("Cta_Mvto", "count(*)", "idCta_Dcto", Cta_Dcto, True))
'                idFormaPago = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", idFormaPago, True)
'
'                If idFormaPago <> "06090001" And idFormaPago <> "06090004" And Pagos <> 0 Then
'                   StrMsgError = "No se puede modificar el estado, porque la factura generada tiene pagos en cuentas por cobrar"
'                   GoTo Err
'                End If
'            End If
'
'        ElseIf txtCod_Documento.Text = "01" Then
'            csql = "select dv.idDocumento, dv.idSerie, dv.idDocVentas, dv.estDocVentas " & _
'                     "from docreferencia dr, docventas dv " & _
'                     "where dr.TipoDocReferencia = dv.idDocumento " & _
'                     "and dr.serieDocReferencia = dv.idSerie " & _
'                     "and dr.numDocReferencia = dv.idDocVentas " & _
'                     "and dr.idEmpresa = dv.idEmpresa " & _
'                     "and dr.TipoDocOrigen = '01' " & _
'                     "and dr.serieDocOrigen = '" & txtserie.Text & "' " & _
'                     "and dr.numDocOrigen = '" & txtnumdoc.Text & "' " & _
'                     "and dr.TipoDocReferencia = '86' " & _
'                     "and dr.idEmpresa = '" & glsEmpresa & "' "
'            If rsg.State = 1 Then rsg.Close
'            rsg.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'
'            If rsg.RecordCount <> 0 Then
'                rsg.MoveFirst
'                '--- Verificamos el estado de la factura
'                If rsg.Fields("estDocVentas") = "ANU" Then
'                    StrMsgError = "La guia relacionada esta anulada, no se puede habilitar la Factura"
'                    GoTo Err
'                End If
'
'                '--- Verificamos si la factura tiene pagos
'                idFormaPago = traerCampo("movcajasdet", "idFormadePago", "iddocventas", txtnumdoc.Text, True, "idSerie = '" & txtserie.Text & "' and idDocumento = '01' and idTipoMovCaja = '99990002'")
'                cabrev = "Fac"
'                cdocumento = cabrev & serieFactura & "/" & numFactura
'                Cta_Dcto = traerCampo("Cta_Dcto", "idCta_Dcto", "Nro_Comp", cdocumento, True)
'                Pagos = Val(traerCampo("Cta_Mvto", "count(*)", "idCta_Dcto", Cta_Dcto, True))
'                idFormaPago = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", idFormaPago, True)
'
'                If idFormaPago <> "06090001" And idFormaPago <> "06090004" And Pagos <> 0 Then
'                   StrMsgError = "No se puede modificar el estado, la factura tiene pagos en cuentas por cobrar"
'                   GoTo Err
'                End If
'            End If
'        End If
        
'        srtEstDoc = traerCampo("Docventas", "estDocventas", "idDocventas", Trim(txtnumdoc.Text), True, "idSerie = '" & Trim(txtserie.Text) & "' And idDocumento = '" & Trim(txtCod_Documento.Text) & "' And idSucursal= '" & glsSucursal & "'")
'
        csql = "UPDATE docventas SET EstDocImportado = 'N' " & _
                 "WHERE iddocventas = '" & txtnumdoc.Text & "' " & _
                 "AND idserie = '" & txtserie.Text & "' " & _
                 "AND iddocumento = '" & txtCod_Documento.Text & "' " & _
                 "AND idEmpresa = '" & glsEmpresa & "' "
        Cn.Execute csql
          
        csql = "UPDATE DocVentasDet SET EstDocImportado = 'N', CantidadImp = 0 " & _
                 "WHERE iddocventas = '" & txtnumdoc.Text & "' " & _
                 "AND idserie = '" & txtserie.Text & "' " & _
                 "AND iddocumento = '" & txtCod_Documento.Text & "' " & _
                 "AND idEmpresa = '" & glsEmpresa & "' "
        Cn.Execute csql
          
'        srtValesCab = Trim("" & traerCampo("DocVentas", "idValesCab", "idDocventas", Trim(txtnumdoc.Text), True, "idSerie = '" & Trim(txtserie.Text) & "' And idDocumento  ='" & Trim(txtCod_Documento.Text) & "' And idSucursal  ='" & glsSucursal & "'"))
'
'        If Len(srtValesCab) > 0 And srtEstDoc = "ANU" Then
'            csql = "UPDATE ValesCab SET estValeCab = 'GEN' " & _
'                     "WHERE idValesCab = '" & srtValesCab & "' " & _
'                     "AND idEmpresa = '" & glsEmpresa & "' " & _
'                     "AND idSucursal = '" & glsSucursal & "' " & _
'                     "AND TipoVale = '" & IIf(txtCod_Documento.Text = "07", "I", "S") & "'  "
'            Cn.Execute csql
'        End If
                
        MsgBox ("Estado Modificado Correctamente."), vbInformation, App.Title
        limpiar
    Else
        StrMsgError = "Proceso Cancelado"
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Documento_Click()
    
    txtGls_Documento.Text = traerCampo("Documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)

End Sub

Private Sub limpiar()
    
    txtCod_Documento.Text = ""
    txtGls_Documento.Text = ""
    txtserie.Text = ""
    txtnumdoc.Text = ""

End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtnumdoc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtnumdoc_LostFocus()
    
    txtnumdoc.Text = Format("" & txtnumdoc.Text, "00000000")

End Sub

Public Sub validar(StrTipo As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodUsuarioAutorizacion As String
Dim IndEvaluacion As Integer
    
    If StrTipo = "A" Then
        IndEvaluacion = 0

'        frmAprobacion.MostrarForm "06", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
'        If StrMsgError <> "" Then GoTo Err
        
'        If IndEvaluacion = 0 Then
'            strCodUsuarioAutorizacion = ""
'            Exit Sub
'        Else
            actualizar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
'    End If
    
Exit Sub
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtserie_LostFocus()
    
    txtserie.Text = Format("" & txtserie.Text, "0000")

End Sub
