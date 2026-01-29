VERSION 5.00
Begin VB.Form FrmProcesaSaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesar Saldos"
   ClientHeight    =   3345
   ClientLeft      =   5730
   ClientTop       =   3135
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   2550
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cbxMesHasta 
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
         ItemData        =   "FrmProcesaSaldos.frx":0000
         Left            =   1665
         List            =   "FrmProcesaSaldos.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   2340
      End
      Begin VB.ComboBox CbxAno 
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
         ItemData        =   "FrmProcesaSaldos.frx":02F6
         Left            =   1665
         List            =   "FrmProcesaSaldos.frx":0312
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1155
         Width           =   2340
      End
      Begin VB.ComboBox CbxMes 
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
         ItemData        =   "FrmProcesaSaldos.frx":0346
         Left            =   1665
         List            =   "FrmProcesaSaldos.frx":036E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1605
         Width           =   2340
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "* No Olvidar correr el proceso hasta el mes Actual*"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   660
         Width           =   4605
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "ADVERTENCIA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   990
         TabIndex        =   10
         Top             =   210
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Height          =   1035
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   5010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   1035
         TabIndex        =   8
         Top             =   2085
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   1035
         TabIndex        =   6
         Top             =   1650
         Width           =   465
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
         Left            =   1035
         TabIndex        =   5
         Top             =   1245
         Width           =   300
      End
   End
   Begin VB.CommandButton BtnProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   930
      TabIndex        =   1
      Top             =   2760
      Width           =   1635
   End
   Begin VB.CommandButton BtnSalir 
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
      Height          =   465
      Left            =   2685
      TabIndex        =   0
      Top             =   2760
      Width           =   1635
   End
End
Attribute VB_Name = "FrmProcesaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp          As New ADODB.Recordset
Dim strParamePPro   As String
Option Explicit

Private Sub BtnProcesar_Click()
Dim dblVVUnit           As Double
Dim dblIGVUnit          As Double
Dim dblPVUnit           As Double
Dim dblTotalVVBruto     As Double
Dim dblTotalPVBruto     As Double
Dim dblTotalVVNeto      As Double
Dim dblTotalIGVNeto     As Double
Dim dblTotalPVNeto      As Double
Dim RsDetalle           As New ADODB.Recordset
Dim strIdValesCab       As String
Dim strIdValesCabIng    As String
Dim strIdValesConver    As String
Dim indTrans            As Boolean
Dim StrMsgError         As String
Dim n As Integer
Dim xFecha              As String
Dim Periodo             As String
On Error GoTo Err
    
    indTrans = False
      
    If Val(right(CbxMes.Text, 2)) > Val(right(cbxMesHasta.Text, 2)) Then
        MsgBox "El mes desde no puede ser mayor al mes hasta", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Year(Now) = Val(cbxAno.Text) And Month(Now) < Val(right(cbxMesHasta.Text, 2)) Then
        MsgBox "El mes hasta no puede ser mayor al mes actual", vbInformation, Me.Caption
        Exit Sub
    End If
            
    If MsgBox("¿Está seguro(a) de realizar el proceso? ", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
         
    Me.MousePointer = vbHourglass
    
    For n = Val(right(CbxMes.Text, 2)) To Val(right(cbxMesHasta.Text, 2))
                                
        xFecha = "01/" & right("00" & Val(n), 2) & "/" & Val(cbxAno.Text)
        Periodo = Format(xFecha, "yyyy-mm-dd")
                
        If Year(Now) = Val(cbxAno.Text) And Month(Now) = Val(n) Then
            Periodo = Format(Now, "yyyy-mm-dd")
        
            csql = "EXEC Spu_RegeneraSaldos '','" & Periodo & "' "
            Cn.Execute csql
        Else
            csql = "EXEC Spu_RegeneraSaldos_Proc '','" & Periodo & "' "
            Cn.Execute csql
        End If
        
    Next n
    
    Me.MousePointer = vbNormal
    
    MsgBox "Fin del proceso", vbInformation, "Aviso de Sistema"
    
Exit Sub
Err:
Me.MousePointer = vbNormal
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, "Aviso de Sistema"
Exit Sub
Resume
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim dblMes        As Double
Dim dblAnnio      As Double
Dim i           As Double
Dim StrMsgError As String
On Error GoTo Err
    
    Me.top = 0
    Me.left = 0
    
    dblMes = Format(getFechaSistema, "mm")
    dblAnnio = Format(Year(getFechaSistema), "0000")
    
    cbxAno.Clear
    For i = 2008 To Val(dblAnnio)
        cbxAno.AddItem i
    Next
    cbxAno.AddItem "Periodo Actual"
    
    CbxMes.Clear
    CbxMes.AddItem "ENERO" & Space(80) & "01"
    CbxMes.AddItem "FEBRERO" & Space(80) & "02"
    CbxMes.AddItem "MARZO" & Space(80) & "03"
    CbxMes.AddItem "ABRIL" & Space(80) & "04"
    CbxMes.AddItem "MAYO" & Space(80) & "05"
    CbxMes.AddItem "JUNIO" & Space(80) & "06"
    CbxMes.AddItem "JULIO" & Space(80) & "07"
    CbxMes.AddItem "AGOSTO" & Space(80) & "08"
    CbxMes.AddItem "SETIEMBRE" & Space(80) & "09"
    CbxMes.AddItem "OCTUBRE" & Space(80) & "10"
    CbxMes.AddItem "NOVIEMBRE" & Space(80) & "11"
    CbxMes.AddItem "DICIEMBRE" & Space(80) & "12"
    
    cbxMesHasta.Clear
    cbxMesHasta.AddItem "ENERO" & Space(80) & "01"
    cbxMesHasta.AddItem "FEBRERO" & Space(80) & "02"
    cbxMesHasta.AddItem "MARZO" & Space(80) & "03"
    cbxMesHasta.AddItem "ABRIL" & Space(80) & "04"
    cbxMesHasta.AddItem "MAYO" & Space(80) & "05"
    cbxMesHasta.AddItem "JUNIO" & Space(80) & "06"
    cbxMesHasta.AddItem "JULIO" & Space(80) & "07"
    cbxMesHasta.AddItem "AGOSTO" & Space(80) & "08"
    cbxMesHasta.AddItem "SETIEMBRE" & Space(80) & "09"
    cbxMesHasta.AddItem "OCTUBRE" & Space(80) & "10"
    cbxMesHasta.AddItem "NOVIEMBRE" & Space(80) & "11"
    cbxMesHasta.AddItem "DICIEMBRE" & Space(80) & "12"

    CbxMes.ListIndex = dblMes - 1
    cbxMesHasta.ListIndex = dblMes - 1
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = dblAnnio Then Exit For
    Next
    
    strParamePPro = traerCampo("Parametros", "ValParametro", "GlsParametro", "EXTRA_PRECIO_PROMEDIO", True)
    
    
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, "Aviso de Sistema"
End Sub

Private Sub ObtenerPrecioPromedio(StrItem As String, StrValeSalida As String, strCodPro As String, strCodAlmOri As String, strCodMO As String, strFecEmi As String, dblTipCamb As Double, intAfecto As Integer, dblVVUnit As Double, dblIGVUnit As Double, dblPVUnit As Double, StrMsgError As String)
Dim rst                 As New ADODB.Recordset
Dim VVUnit              As Double
Dim IGVUnit             As Double
Dim PVUnit              As Double
Dim TipoCambio          As Double
On Error GoTo Err
    
    TipoCambio = Val(Format(dblTipCamb, "0.000"))
    
    If strParamePPro = "0" Then
    
       VVUnit = 0#
       IGVUnit = 0#
       PVUnit = 0#
       
       csql = "select a.idmoneda,b.VVUnit, b.IGVUnit, b.PVUnit from docventas a " & _
                "inner join docventasdet b " & _
                "on a.iddocventas = b.iddocventas " & _
                "and a.idserie = b.idserie " & _
                "and a.iddocumento = b.iddocumento " & _
                "and a.idempresa = b.idempresa " & _
                "where a.idempresa = '" & glsEmpresa & "' " & _
                "and a.iddocumento = '94' " & _
                "and idproducto = '" & strCodPro & "' " & _
                "order by a.fecemision desc   Limit 1 "
    
       If rst.State = 1 Then rst.Close: Set rst = Nothing
       rst.Open csql, Cn, adOpenStatic, adLockReadOnly
       If Not rst.EOF Then
               
            If strCodMO = "PEN" Then
                If Trim("" & rst.Fields("idMoneda")) = "PEN" Then
                    VVUnit = Val(Format(rst.Fields("VVUnit"), "0.00"))
                    IGVUnit = Val(Format(rst.Fields("IGVUnit"), "0.00"))
                    PVUnit = Val(Format(rst.Fields("PVUnit"), "0.00"))
                Else
                    VVUnit = Val(Format((Val(Format(rst.Fields("VVUnit"), "0.00")) * TipoCambio), "0.00"))
                    IGVUnit = Val(Format((Val(Format(rst.Fields("IGVUnit"), "0.00")) * TipoCambio), "0.00"))
                    PVUnit = Val(Format((Val(Format(rst.Fields("PVUnit"), "0.00")) * TipoCambio), "0.00"))
                End If
            Else
                If Trim("" & rst.Fields("idMoneda")) = "USD" Then
                    VVUnit = Val(Format(rst.Fields("VVUnit"), "0.00"))
                    IGVUnit = Val(Format(rst.Fields("IGVUnit"), "0.00"))
                    PVUnit = Val(Format(rst.Fields("PVUnit"), "0.00"))
                Else
                    VVUnit = Val(Format((Val(Format(rst.Fields("VVUnit"), "0.00")) / TipoCambio), "0.00"))
                    IGVUnit = Val(Format((Val(Format(rst.Fields("IGVUnit"), "0.00")) / TipoCambio), "0.00"))
                    PVUnit = Val(Format((Val(Format(rst.Fields("PVUnit"), "0.00")) / TipoCambio), "0.00"))
                End If
            End If
            
            dblVVUnit = Val(Format(VVUnit, "0.00"))
            dblIGVUnit = Val(Format(IGVUnit, "0.00"))
            dblPVUnit = Val(Format(PVUnit, "0.00"))
                
       Else
       
             VVUnit = 0#
             IGVUnit = 0#
             PVUnit = 0#
        
             VVUnit = traerCostoUnit(StrItem, StrValeSalida, strCodPro, strCodAlmOri, Format(strFecEmi, "yyyy-mm-dd"), strCodMO, StrMsgError)
             If StrMsgError <> "" Then GoTo Err
             
             If Val(left(VVUnit, 4)) < 0 Then
                VVUnit = 0#
                IGVUnit = 0#
                PVUnit = 0#
             
                dblVVUnit = Val(Format(VVUnit, "0.00"))
                dblIGVUnit = Val(Format(IGVUnit, "0.00"))
                dblPVUnit = Val(Format(PVUnit, "0.00"))
             Else
                
                VVUnit = Val(Format(VVUnit, "#,###.0000"))
             
                procesaMoneda strCodMO, strCodMO, 0, VVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit, TipoCambio, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
                dblVVUnit = Val(Format(dblVVUnit, "0.00"))
                dblIGVUnit = Val(Format(dblIGVUnit, "0.00"))
                dblPVUnit = Val(Format(dblPVUnit, "0.00"))
                   
             End If
             
       End If
       
    Else
        'Extrae Precio Procio Promedio
        VVUnit = traerCostoUnit(StrItem, StrValeSalida, strCodPro, strCodAlmOri, Format(strFecEmi, "yyyy-mm-dd"), strCodMO, StrMsgError)
        If StrMsgError <> "" Then GoTo Err
        
        If Val(left(VVUnit, 4)) < 0 Then
           
           VVUnit = 0#
           IGVUnit = 0#
           PVUnit = 0#
        
           dblVVUnit = Val(Format(VVUnit, "0.00"))
           dblIGVUnit = Val(Format(IGVUnit, "0.00"))
           dblPVUnit = Val(Format(PVUnit, "0.00"))
           
        Else
           
            VVUnit = Val(Format(VVUnit, "#,###.0000"))
                 
            procesaMoneda strCodMO, strCodMO, 0, VVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit, TipoCambio, StrMsgError
            If StrMsgError <> "" Then GoTo Err
                 
            dblVVUnit = Val(Format(dblVVUnit, "0.00"))
            dblIGVUnit = Val(Format(dblIGVUnit, "0.00"))
            dblPVUnit = Val(Format(dblPVUnit, "0.00"))
        
        End If
    End If
    
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function traerCostoUnit(ByVal StrItem As String, ByVal StrValeSalida As String, ByVal codproducto As String, ByVal codalmacen As String, ByVal PFecha As String, ByVal CodMoneda As String, ByRef StrMsgError As String) As Double
Dim CosUni  As ADODB.Recordset
On Error GoTo Err

    
   'csql = "Select(SUM(IF((valescab.tipoVale = 'I'),(valesdet.Cantidad),((valesdet.Cantidad) * -(1))) * " & _
   ' "CASE '" & CodMoneda & "' " & _
   ' "WHEN 'PEN' THEN IF(valescab.idMoneda = 'PEN', valesdet.VVUnit,valesdet.VVUnit * ValesCab.TipoCambio) " & _
   ' "WHEN 'USD' THEN IF(valescab.idMoneda = 'USD', valesdet.VVUnit,valesdet.VVUnit / ValesCab.TipoCambio) " & _
   ' "End) / " & _
   ' "SUM(IF((valescab.tipoVale = 'I'),(valesdet.Cantidad),((valesdet.Cantidad) * -(1))))) AS COSTO_UNITARIO "
   ' csql = csql & "FROM valescab " & _
   '  "INNER JOIN valesdet  " & _
   '     "ON valescab.idValesCab = valesdet.idValesCab  " & _
   '     "AND valescab.idEmpresa = valesdet.idEmpresa  " & _
   '     "AND valescab.idSucursal = valesdet.idSucursal  " & _
   '     "AND valescab.tipoVale = valesdet.tipoVale " & _
   '   "INNER JOIN conceptos  " & _
   '     "ON valescab.idConcepto = conceptos.idConcepto  " & _
   '   "LEFT JOIN tiposdecambio t " & _
   '     "ON valescab.fechaEmision = t.fecha "
   ' csql = csql & "WHERE "
   ' csql = csql & "valescab.idEmpresa = '" & glsEmpresa & "' "
   ' csql = csql & "AND (valescab.idPeriodoInv) IN " & _
   '                 "(" & _
   '                     "SELECT pi.idPeriodoInv " & _
   '                     "FROM periodosinv pi " & _
   '                     "WHERE pi.idEmpresa = valescab.idEmpresa AND pi.idSucursal = valescab.idSucursal and pi.FecInicio <= '" & Format(PFecha, "yyyy-mm-dd") & "' " & _
   '                     "and (pi.FecFin >= '" & Format(PFecha, "yyyy-mm-dd") & "' or pi.FecFin is null)" & _
   '                 ") "
   ' csql = csql & " AND valescab.fechaEmision <= '" & PFecha & "' And valesdet.idProducto = '" & codproducto & "' "
   ' csql = csql & "AND valescab.idAlmacen = '" & codalmacen & "' "
   ' csql = csql & "AND valescab.estValeCab <> 'ANU' "'

    csql = "SELECT VVUnit FROM ValesDet " & _
           "Where idValesCab = '" & Trim("" & StrValeSalida) & "' And Item = " & StrItem & " " & _
           "And idEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' AND TipoVale  = 'S'"
                
                
    Set CosUni = New ADODB.Recordset
    CosUni.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not CosUni.EOF Then
       'traerCostoUnit = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
       traerCostoUnit = IIf(IsNull(CosUni.Fields("VVUnit")), 0, CosUni.Fields("VVUnit"))
    End If
    If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    
    Exit Function
    
Err:
If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Private Sub procesaMoneda(strMonProd As String, strMonDoc As String, intTipoValor As Integer, dblValor As Double, intAfecto As Integer, ByRef dblVVUnit As Double, ByRef dblIGVUnit As Double, ByRef dblPVUnit As Double, TipoCambio As Double, StrMsgError As String)
Dim dblIGV  As Double
Dim dblTC   As Double
On Error GoTo Err

    dblIGV = glsIGV
    dblTC = TipoCambio
    If intAfecto = 0 Then dblIGV = 0
    
    If strMonDoc = "USD" Then 'dolares
        If strMonProd = "PEN" Then 'soles
            dblValor = dblValor / dblTC
        End If
    Else 'soles
        If strMonProd = "USD" Then 'dolares
            dblValor = dblValor * dblTC
        End If
    End If
    
    If intTipoValor = 0 Then 'valor venta
        dblVVUnit = dblValor
        dblIGVUnit = dblValor * dblIGV
        dblPVUnit = dblVVUnit + dblIGVUnit
    Else 'precio venta
        dblVVUnit = dblValor / (dblIGV + 1)
        dblIGVUnit = dblValor - dblVVUnit
        dblPVUnit = dblValor
    End If
    
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
End Sub

Private Sub ActualizaIngreso(strIdValesSal As String, strIdValesIng As String, strIdValesConver As String, StrMsgError As String)
Dim TotalOrigen     As Double
Dim TotalCosteo     As Double
Dim TotCantidad     As Double
Dim dblPVUnit       As Double
Dim dblVVUnit       As Double
Dim dblIGVUnit      As Double
Dim dblTotalVVBruto As Double
Dim dblTotalPVBruto As Double
Dim dblTotalVVNeto  As Double
Dim dblTotalIGVNeto As Double
Dim dblTotalPVNeto  As Double
Dim dblCantidad     As Double
On Error GoTo Err
    
    TotalOrigen = 0
    TotalCosteo = 0
    dblCantidad = 0
    
    csql = "Select Sum(TotalPVNeto) As TotalPVNeto From ValesDet V " & _
           "Where V.idValesCab = '" & strIdValesSal & "' " & _
           "AND V.idEmpresa = '" & glsEmpresa & "' AND V.idSucursal = '" & glsSucursal & "' AND V.TipoVale  = 'S'"
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    rstemp.Open csql, Cn, adOpenDynamic, adLockReadOnly
    If Not rstemp.EOF Then
        TotalOrigen = Val("" & rstemp.Fields("TotalPVNeto").Value)
    End If
    
    csql = "Select Sum(TotalPVNeto) As TotalPVNeto From ValesConverCosteo V " & _
           "Where V.idValesConver = '" & strIdValesConver & "' " & _
           "AND V.idEmpresa = '" & glsEmpresa & "' AND V.idSucursal = '" & glsSucursal & "' "
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    rstemp.Open csql, Cn, adOpenDynamic, adLockReadOnly
    If Not rstemp.EOF Then
        TotalCosteo = Val("" & rstemp.Fields("TotalPVNeto").Value)
    End If

    csql = "Select Sum(Cantidad) As Cantidad From ValesDet V " & _
           "Where V.idValesCab = '" & strIdValesIng & "' " & _
           "AND V.idEmpresa = '" & glsEmpresa & "' AND V.idSucursal = '" & glsSucursal & "' AND V.TipoVale  = 'I'"
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    rstemp.Open csql, Cn, adOpenDynamic, adLockReadOnly
    If Not rstemp.EOF Then
        TotCantidad = Val("" & rstemp.Fields("Cantidad").Value)
    End If
    
    
    csql = "Select Cantidad, Afecto, Item From ValesDet V " & _
           "Where V.idValesCab = '" & strIdValesIng & "' " & _
           "AND V.idEmpresa = '" & glsEmpresa & "' AND V.idSucursal = '" & glsSucursal & "' AND V.TipoVale  = 'I' " & _
           "Order By V.Item"
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    rstemp.Open csql, Cn, adOpenDynamic, adLockReadOnly
    If Not rstemp.EOF Then
        Do While Not rstemp.EOF
            
            dblPVUnit = Val((((((TotalOrigen + TotalCosteo) / TotCantidad) * rstemp.Fields("Cantidad").Value) / rstemp.Fields("Cantidad").Value)))
            dblVVUnit = Val(dblPVUnit / (1 + glsIGV))
            dblIGVUnit = Val(dblVVUnit * glsIGV)
            
            dblTotalVVBruto = rstemp.Fields("Cantidad").Value * dblVVUnit
            dblTotalPVBruto = rstemp.Fields("Cantidad").Value * dblPVUnit
            
            dblTotalVVNeto = dblTotalVVBruto '- dblDctoVV
            If rstemp.Fields("Afecto").Value = 1 Then
                dblTotalIGVNeto = dblTotalVVNeto * glsIGV
            Else
                dblTotalIGVNeto = 0
            End If
            dblTotalPVNeto = dblTotalVVNeto + dblTotalIGVNeto
             
            csql = "Update ValesDet Set VVUnit = " & Format(dblVVUnit, "0.00") & ", IGVUnit = " & Format(dblIGVUnit, "0.00") & ", PVUnit = " & Format(dblPVUnit, "0.00") & ", " & _
                   "TotalVVNeto = " & Format(dblTotalVVNeto, "0.00") & ", TotalIGVNeto =  " & Format(dblTotalIGVNeto, "0.00") & ", TotalPVNeto = " & Format(dblTotalPVNeto, "0.00") & " " & _
                   "Where idValesCab = '" & strIdValesIng & "' And Item = " & rstemp.Fields("Item").Value & " " & _
                   "And idEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' AND TipoVale  = 'I'"
            Cn.Execute (csql)
             
             
            rstemp.MoveNext
        Loop
        
        'Actualizamos cabecera del vale ingreso
        csql = "Select Sum(V.VVUnit * V.Cantidad) As ValorTotal " & _
                "From ValesDet V " & _
                "Where V.idValesCab = '" & strIdValesIng & "' " & _
                "AND V.idEmpresa = '" & glsEmpresa & "' AND V.idSucursal = '" & glsSucursal & "' AND V.TipoVale  = 'I'"
        If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
        rstemp.Open csql, Cn, adOpenDynamic, adLockReadOnly
        If Not rstemp.EOF Then
        csql = "Update ValesCab Set ValorTotal = " & rstemp.Fields("ValorTotal").Value & ", IgvTotal = " & rstemp.Fields("ValorTotal").Value * glsIGV & ", " & _
                "PrecioTotal = " & rstemp.Fields("ValorTotal").Value + (rstemp.Fields("ValorTotal").Value * glsIGV) & " " & _
                "Where idValesCab = " & strIdValesIng & " " & _
                "AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND TipoVale  = 'I'"
                Cn.Execute (csql)
        End If
    End If
    
     
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
Exit Sub
Resume
End Sub











