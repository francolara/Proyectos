VERSION 5.00
Begin VB.Form frmAsientoContableConsistenciaDeCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consistencias De Cuentas"
   ClientHeight    =   4500
   ClientLeft      =   5400
   ClientTop       =   3000
   ClientWidth     =   4005
   LinkTopic       =   "Generar Asientos Contables"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4005
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
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3915
      Width           =   1140
   End
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
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3915
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   3750
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   540
         TabIndex        =   3
         Top             =   1080
         Width           =   2805
         Begin VB.CheckBox ChkCentro 
            Caption         =   "Centro de Costo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   12
            Top             =   1620
            Width           =   1950
         End
         Begin VB.CheckBox ChkReferencia 
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   11
            Top             =   720
            Width           =   1950
         End
         Begin VB.CheckBox ChkNoExiste 
            Caption         =   "No existe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   10
            Top             =   1170
            Width           =   1950
         End
         Begin VB.CheckBox ChkRuc 
            Caption         =   "Código Auxiliar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   9
            Top             =   2070
            Width           =   1950
         End
         Begin VB.CheckBox ChkNoDetalle 
            Caption         =   "No es de Detalle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   8
            Top             =   270
            Width           =   1950
         End
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
         ItemData        =   "frmAsientoContableConsistenciaDeCuentas.frx":0000
         Left            =   1170
         List            =   "frmAsientoContableConsistenciaDeCuentas.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   1080
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
         ItemData        =   "frmAsientoContableConsistenciaDeCuentas.frx":0050
         Left            =   1170
         List            =   "frmAsientoContableConsistenciaDeCuentas.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   1845
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
         Left            =   585
         TabIndex        =   5
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label3 
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
         Left            =   585
         TabIndex        =   4
         Top             =   315
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmAsientoContableConsistenciaDeCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim csql            As String
Dim TbTabla         As New ADODB.Recordset
Dim linea           As String
Dim strcnConta      As String
Dim cadena          As String
Dim strAnno         As String

Private Sub cmbCancelar_Click()
    
    Unload Me
    
End Sub

Private Sub cmbOperar_Click()
Dim StrMsgError     As String
Dim cCompro         As String
Dim xOrigen         As String
Dim xFecha          As Variant
Dim nDebes          As Double, nHabers As Double
Dim nDebeD          As Double, nHaberD As Double
Dim RsCuadrar           As New ADODB.Recordset
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim GlsReporte      As String
Dim GlsForm         As String
Dim CnConta         As New ADODB.Connection
On Error GoTo Err
    
    strcnConta = "dsn=dnsContabilidad"
    
    CnConta.CursorLocation = adUseClient
    CnConta.Open strcnConta
            
    If rsAsientosContables.State = 1 Then
        If rsAsientosContables.RecordCount <> 0 Then
            If RsCuadrar.State = 1 Then RsCuadrar.Close
        
            RsCuadrar.Fields.Append "idComprobante", adVarChar, 9, adFldRowID
            RsCuadrar.Fields.Append "FecCompro", adVarChar, 30, adFldRowID
            RsCuadrar.Fields.Append "GlsEmpresa", adVarChar, 250, adFldIsNullable
            RsCuadrar.Fields.Append "GlsDetalle", adVarChar, 250, adFldIsNullable
            RsCuadrar.Fields.Append "GlsDescripcion", adVarChar, 250, adFldIsNullable
            RsCuadrar.Fields.Append "ValDebe", adDouble, adFldRowID
            RsCuadrar.Fields.Append "ValHaber", adDouble, adFldIsNullable
            RsCuadrar.Fields.Append "GlsTitulo", adVarChar, 250, adFldRowID
            RsCuadrar.Fields.Append "idCtaContable", adVarChar, 150, adFldRowID
            RsCuadrar.Open
            
            rsAsientosContables.MoveFirst
            If Trim(rsAsientosContables.Fields("IdPeriodo")) = Trim(cbxAno.Text & Format(cbxMes.ListIndex + 1, "00")) Then
                Do While Not rsAsientosContables.EOF
                    If ChkNoDetalle.Value = 1 Then
                        If nodetalle(rsAsientosContables.Fields("IdCtaContable") & "") = False Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("GlsDescripcion") = "No es de detalle"
                            RsCuadrar.Fields("idComprobante") = rsAsientosContables.Fields("idComprobante") & ""
                            RsCuadrar.Fields("FecCompro") = rsAsientosContables.Fields("FecCompro")
                            RsCuadrar.Fields("idCtaContable") = rsAsientosContables.Fields("IdCtaContable") & ""
                            RsCuadrar.Fields("GlsDetalle") = rsAsientosContables.Fields("GlsDetalle") & ""
                            If rsAsientosContables.Fields("IdTipoDH") = "D" Then
                                RsCuadrar.Fields("ValDebe") = rsAsientosContables.Fields("TotalImporteS")
                            Else
                                RsCuadrar.Fields("ValHaber") = rsAsientosContables.Fields("TotalImporteS")
                            End If
                            RsCuadrar.Fields("GlsTitulo") = "CONSISTENCIA DE COMPROBANTES DEL MES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2)))
                            RsCuadrar.Update
                        End If
                    End If
                    If ChkReferencia.Value = 1 Then
                        If val_referere(rsAsientosContables.Fields("idCtaContable") & "", rsAsientosContables.Fields("NumReferencia") & "") = True Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("GlsDescripcion") = "Necesita referencia"
                            RsCuadrar.Fields("idComprobante") = rsAsientosContables.Fields("idComprobante") & ""
                            RsCuadrar.Fields("FecCompro") = rsAsientosContables.Fields("FecCompro")
                            RsCuadrar.Fields("idCtaContable") = rsAsientosContables.Fields("idCtaContable") & ""
                            RsCuadrar.Fields("GlsDetalle") = rsAsientosContables.Fields("GlsDetalle") & ""
                            If rsAsientosContables.Fields("IdTipoDH") = "D" Then
                                RsCuadrar.Fields("ValDebe") = rsAsientosContables.Fields("TotalImporteS")
                            Else
                                RsCuadrar.Fields("ValHaber") = rsAsientosContables.Fields("TotalImporteS")
                            End If
                            RsCuadrar.Fields("GlsTitulo") = "CONSISTENCIA DE COMPROBANTES DEL MES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2)))
                            RsCuadrar.Update
                        End If
                    End If
                    If ChkNoExiste.Value = 1 Then
                        linea = Space(132)
                        If noexiste("" & rsAsientosContables.Fields("idCtaContable")) & "" = False Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("GlsDescripcion") = "No existe"
                            RsCuadrar.Fields("idComprobante") = rsAsientosContables.Fields("idComprobante") & ""
                            RsCuadrar.Fields("FecCompro") = rsAsientosContables.Fields("FecCompro")
                            RsCuadrar.Fields("idCtaContable") = rsAsientosContables.Fields("idCtaContable") & ""
                            RsCuadrar.Fields("GlsDetalle") = rsAsientosContables.Fields("GlsDetalle") & ""
                            If rsAsientosContables.Fields("IdTipoDH") = "D" Then
                                RsCuadrar.Fields("ValDebe") = rsAsientosContables.Fields("TotalImporteS")
                            Else
                                RsCuadrar.Fields("ValHaber") = rsAsientosContables.Fields("TotalImporteS")
                            End If
                            RsCuadrar.Fields("GlsTitulo") = "CONSISTENCIA DE COMPROBANTES DEL MES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2)))
                            RsCuadrar.Update
                        End If
                    End If
                    If ChkCentro.Value = 1 Then
                        If val_centro("" & rsAsientosContables.Fields("idCtaContable") & "", rsAsientosContables.Fields("IdCosto") & "") = False Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("GlsDescripcion") = "Necesita centro de costo"
                            RsCuadrar.Fields("idComprobante") = rsAsientosContables.Fields("idComprobante") & ""
                            RsCuadrar.Fields("FecCompro") = rsAsientosContables.Fields("FecCompro")
                            RsCuadrar.Fields("idCtaContable") = rsAsientosContables.Fields("idCtaContable") & ""
                            RsCuadrar.Fields("GlsDetalle") = rsAsientosContables.Fields("GlsDetalle") & ""
                            If rsAsientosContables.Fields("IdTipoDH") = "D" Then
                                RsCuadrar.Fields("ValDebe") = rsAsientosContables.Fields("TotalImporteS")
                            Else
                                RsCuadrar.Fields("ValHaber") = rsAsientosContables.Fields("TotalImporteS")
                            End If
                            RsCuadrar.Fields("GlsTitulo") = "CONSISTENCIA DE COMPROBANTES DEL MES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2)))
                            RsCuadrar.Update
                        End If
                    End If
                    If ChkRuc.Value = 1 Then
                      linea = Space(132)
                       If val_codauxiliar(rsAsientosContables.Fields("idCtaContable") & "", rsAsientosContables.Fields("CtaAuxiliar") & "") = False Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("GlsDescripcion") = "Necesita código auxiliar"
                            RsCuadrar.Fields("idComprobante") = rsAsientosContables.Fields("idComprobante") & ""
                            RsCuadrar.Fields("FecCompro") = rsAsientosContables.Fields("FecCompro")
                            RsCuadrar.Fields("idCtaContable") = rsAsientosContables.Fields("idCtaContable") & ""
                            RsCuadrar.Fields("GlsDetalle") = rsAsientosContables.Fields("GlsDetalle") & ""
                            If rsAsientosContables.Fields("IdTipoDH") = "D" Then
                                RsCuadrar.Fields("ValDebe") = rsAsientosContables.Fields("TotalImporteS")
                            Else
                                RsCuadrar.Fields("ValHaber") = rsAsientosContables.Fields("TotalImporteS")
                            End If
                            
                            RsCuadrar.Fields("GlsTitulo") = "CONSISTENCIA DE COMPROBANTES DEL MES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2)))
                            RsCuadrar.Update
                       End If
                    End If
                    rsAsientosContables.MoveNext
                Loop
    
                Screen.MousePointer = 11
            
                GlsReporte = "RptConsisDeCuentas.rpt"
                GlsForm = "Consistencias de Cuentas"
                Set reporte = aplicacion.OpenReport(gStrRutaRpts & GlsReporte)
                If Not RsCuadrar.EOF And Not RsCuadrar.BOF Then
                    reporte.Database.SetDataSource RsCuadrar, 3
                    vistaPrevia.CRViewer91.ReportSource = reporte
                    vistaPrevia.CRViewer91.ViewReport
                    vistaPrevia.CRViewer91.DisplayGroupTree = False
                    Screen.MousePointer = 0
                    vistaPrevia.WindowState = 2
                    vistaPrevia.Show
                Else
                    Screen.MousePointer = 0
                    MsgBox "No hay inconsistencias.", vbInformation, App.Title
                End If
            Else
                MsgBox "El mes no corresponde con los movimientos generados.", vbInformation, App.Title
            End If
        Else
            MsgBox "No existen movimientos generados para la consistencia.", vbInformation, App.Title
        End If
    Else
         MsgBox "No existen movimientos generados para la consistencia.", vbInformation, App.Title
    End If
    
    cmbCancelar.SetFocus
    
    CnConta.Close
    
    Exit Sub
    
Err:
    Screen.MousePointer = 0
        
    If TypeName(RsCuadrar) = "Recordset" Then
        If RsCuadrar.State = 1 Then RsCuadrar.Close
        Set RsCuadrar = Nothing
    End If
    CnConta.Close
    MsgBox Err.Description, vbInformation, App.Title
    
    Exit Sub
    
End Sub

Private Sub Form_Load()
Dim fecha As Date
Dim i As Integer
    
    Me.Height = 5070
    Me.Width = 4245
    
    fecha = Format(getFechaSistema, "dd/mm/yyyy")
    
    strAno = Format(Year(fecha), "0000")
    strMes = Format(Month(fecha), "00")
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = Format(Year(fecha), "0000") Then Exit For
    Next
    
    cbxMes.ListIndex = Val(Format(Month(fecha), "00")) - 1
    
    ChkNoDetalle.Value = 1
    ChkReferencia.Value = 1
    ChkNoExiste.Value = 1
    ChkCentro.Value = 1
    ChkRuc.Value = 1
    
End Sub

Private Function val_centro(cCuenta As String, cCentro As String)
    strAnno = cbxAno.Text
    cadena = IIf(strAnno = "2010", " And IdAnno In('2010')", " And IdAnno Not In('2010')")
    
    csql = "Select IndCCosto From PlanCuentas Where IdCtaContable = '" & cCuenta & "' And IdEmpresa='" & glsEmpresa & "'" & cadena
    TbTabla.Open csql, strcnConta, adOpenStatic, adLockReadOnly
    If Not TbTabla.EOF Then
        If TbTabla.Fields("IndCCosto") = 1 And Trim(cCentro) = "" Then
            val_centro = False
        Else
            val_centro = True
        End If
    End If
       
    TbTabla.Close: Set TbTabla = Nothing
    
End Function

Public Function val_codauxiliar(cCuenta As String, pCodAux As String)
On Error GoTo ERROR
    strAnno = cbxAno.Text
    cadena = IIf(strAnno = "2010", " And IdAnno In('2010')", " And IdAnno Not In('2010')")
    
    csql = "Select Tipo From PlanCuentas Where IdCtaContable = '" & cCuenta & "' And IdEmpresa='" & glsEmpresa & "'" & cadena
    TbTabla.Open csql, strcnConta, adOpenStatic, adLockReadOnly
    If Not TbTabla.EOF Then
        If TbTabla.Fields("Tipo") & "" = "C" Or TbTabla.Fields("Tipo") & "" = "P" Or TbTabla.Fields("Tipo") & "" = "E" Then
            If Len(Trim(pCodAux)) = 0 Then
                val_codauxiliar = False
            Else
                val_codauxiliar = True
            End If
        Else
            val_codauxiliar = True
        End If
    End If
    
    TbTabla.Close: Set TbTabla = Nothing
    
    Exit Function
   
ERROR:
   MsgBox "Se ha producido el sgte. error : " & Err.Description, vbCritical, App.Title
   Exit Function

End Function

Private Function noexiste(cCuenta As String)
    strAnno = cbxAno.Text
    cadena = IIf(strAnno = "2010", " And IdAnno In('2010')", " And IdAnno Not In('2010')")
    
    csql = "Select IdCtaContable From PlanCuentas Where IdCtaContable = '" & cCuenta & "' And IdEmpresa='" & glsEmpresa & "'" & cadena
    TbTabla.Open csql, strcnConta, adOpenStatic, adLockReadOnly
    If Not TbTabla.EOF Then
        noexiste = True
    Else
        noexiste = False
    End If
    TbTabla.Close: Set TbTabla = Nothing
    
End Function

Private Function nodetalle(cCuenta As String)
Dim nGrado As Integer, CtaAux As String
    strAnno = cbxAno.Text
    cadena = IIf(strAnno = "2010", " And IdAnno In('2010')", " And IdAnno Not In('2010')")
    
    csql = "Select GradoCuenta From PlanCuentas Where IdCtaContable = '" & cCuenta & "' And IdEmpresa='" & glsEmpresa & "'" & cadena
    If TbTabla.State = 1 Then TbTabla.Close
    TbTabla.Open csql, strcnConta, adOpenStatic, adLockReadOnly
    
    If Not TbTabla.EOF Then
        nGrado = TbTabla.Fields("GradoCuenta")
        TbTabla.MoveNext
        If TbTabla.EOF = False Then
            CtaAux = Mid(TbTabla.Fields("GradoCuenta") & Space(12 - Len(TbTabla.Fields("GradoCuenta") & "")), 1, Len(Trim(cCuenta)))
            If Trim(cCuenta) = CtaAux And TbTabla.Fields("GradoCuenta") > nGrado Then
                nodetalle = False
            Else
                nodetalle = True
            End If
        Else
            nodetalle = True
        End If
    End If
    
    TbTabla.Close: Set TbTabla = Nothing
    
End Function

Private Function val_referere(cCuenta As String, nRefer As String)
    strAnno = cbxAno.Text
    cadena = IIf(strAnno = "2010", " And IdAnno In('2010')", " And IdAnno Not In('2010')")
    
    csql = "Select IndReferencia From PlanCuentas Where IdCtaContable = '" & cCuenta & "' And IdEmpresa='" & glsEmpresa & "'" & cadena
    TbTabla.Open csql, strcnConta, adOpenStatic, adLockReadOnly
    If Not TbTabla.EOF Then
        If TbTabla.Fields("IndReferencia") = 1 Then
            If Len(Trim(nRefer)) = 0 Then
                val_referere = True
            Else
                val_referere = False
            End If
        Else
            val_referere = False
        End If
    End If
   
    TbTabla.Close: Set TbTabla = Nothing
    
End Function
