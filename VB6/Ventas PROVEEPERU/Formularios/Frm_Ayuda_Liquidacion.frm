VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Frm_Ayuda_Liquidacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Liquidaciones"
   ClientHeight    =   4170
   ClientLeft      =   4245
   ClientTop       =   2610
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Appearance      =   0  'Flat
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
      Left            =   8055
      TabIndex        =   2
      Top             =   3735
      Width           =   1365
   End
   Begin VB.Frame fraDetalle 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   3645
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   9420
      Begin DXDBGRIDLibCtl.dxDBGrid gLiquidaciones 
         Height          =   3375
         Left            =   45
         OleObjectBlob   =   "Frm_Ayuda_Liquidacion.frx":0000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   135
         Width           =   9345
      End
   End
End
Attribute VB_Name = "Frm_Ayuda_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indResultado As Boolean
Dim rsg As New ADODB.Recordset

Private Sub cmdaceptar_Click()
    
    indResultado = True
    Unload Me

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLiquidaciones, True, True, False, False
    indResultado = False
    mostrarLiquidaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarLiquidaciones(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim item As Integer

    item = 0
    csql = "Select dd.Item,d.IdLiquidacion,IdUPP,idCamal,FecRegistro,p.Glspersona,u.DescUnidad From docventasliqcab d " & _
           "Inner Join  docventasliqdet dd  " & _
           "On d.IdLiquidacion = dd.IdLiquidacion And d.IdEmpresa = dd.IdEmpresa And d.IdSucursal = dd.IdSucursal " & _
           "Inner Join UnidadProduccion u  On dd.IdUPP = u.CodUnidProd And dd.idEmpresa = u.idEmpresa " & _
           "Inner Join Personas p On   p.idPersona = d.idCamal  Group By d.IdLiquidacion  Order By d.IdLiquidacion "
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
      
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "CHK", adVarChar, 5, adFldIsNullable
    rsg.Fields.Append "IdLiquidacion", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdUPP", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "idCamal", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "FechaLiq", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "glsGranjaOri", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "glsCamal", adVarChar, 200, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("CHK") = "N"
        rsg.Fields("IdLiquidacion") = ""
        rsg.Fields("IdUPP") = ""
        rsg.Fields("idCamal") = ""
        rsg.Fields("FechaLiq") = ""
        rsg.Fields("glsGranjaOri") = ""
        rsg.Fields("glsCamal") = ""
    Else
        Do While Not rst.EOF
         item = item + 1
            rsg.AddNew
            rsg.Fields("Item") = item
            rsg.Fields("IdLiquidacion") = Trim("" & rst.Fields("IdLiquidacion"))
            rsg.Fields("IdUPP") = Trim("" & rst.Fields("IdUPP"))
            rsg.Fields("idCamal") = Trim("" & rst.Fields("idCamal"))
            rsg.Fields("FechaLiq") = Trim("" & Format(rst.Fields("FecRegistro"), "dd/mm/yyyy"))
            rsg.Fields("glsCamal") = Trim("" & rst.Fields("Glspersona"))
            rsg.Fields("glsGranjaOri") = Trim("" & rst.Fields("DescUnidad"))
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gLiquidaciones, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub MostrarForm(ByRef rsRpta As ADODB.Recordset, ByRef indEvaluaEstado As Boolean, ByRef StrMsgError As String)
On Error GoTo Err
    
    Me.left = 1905
    Me.Show 1
    If indResultado Then
        Set rsRpta = rsg
        indEvaluaEstado = True
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLiquidaciones_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If gLiquidaciones.Dataset.State = dsEdit Then
        gLiquidaciones.Dataset.Post
    End If
    
End Sub
