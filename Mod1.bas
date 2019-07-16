Attribute VB_Name = "Módulo1"

Public Sub dbSECInstance()
    dbSEC.Instance rowSEC, trowSEC, colSEC, naSheetSEC
    'dbSEC.PrintValues
End Sub

Public Sub dbOCUInstance()
    dbOCU.Instance rowOCU, trowOCU, colOCU, naSheetOCU
    'dbOCU.PrintValues
End Sub

Public Sub dbTRAInstance()
    dbTRA.Instance rowTRA, trowTRA, colTRA, naSheetTRA
    'dbTRA.PrintValues
End Sub

Public Sub dbOBRInstance()
    dbOBR.Instance rowOBR, trowOBR, colOBR, naSheetOBR
    'dbOBR.PrintValues
End Sub

Public Sub dbROLInstance()
    dbROL.Instance rowROL, trowROL, colROL, naSheetROL
    'dbROL.PrintValues
End Sub




Public Function Primerdia(inrange As Range)
   Primerdia = 0
   esprimero = False
   For i = 1 To inrange.Columns.count
    If inrange.Cells(1, i) <> "" Then
      If esprimero = False Then
      Primerdia = i
      esprimero = True
      End If
    End If
   Next
   If Primerdia < 0.5 Then
  Primerdia = Int(Primerdia / 2)
  Else
  Primerdia = Int(Primerdia / 2) + 1
  End If
End Function

Public Function Ultimodia(inrange As Range)
   Ultimodia = 0
   esprimero = False
   For i = 1 To inrange.Columns.count
    If inrange.Cells(1, i) <> "" Then
      Ultimodia = i
    End If
   Next
   If Ultimodia < 0.5 Then
  Ultimodia = Int(Ultimodia / 2)
  Else
  Ultimodia = Int(Ultimodia / 2)
  End If
End Function
Public Function Diastrabajables(diasl As Range, prdiatodos As Integer, uldiatodos As Integer, prdiaind As Integer, uldiaind As Integer)
   If (IsNumeric(prdiaind) And Not prdiaind = 0) Then
   prdia = prdiaind
   Else
   prdia = prdiatodos
   End If
    If (IsNumeric(uldiaind) And Not uldiaind = 0) Then
   uldia = uldiaind
   Else
   uldia = uldiatodos
   End If
   For i = prdia To uldia
    diacont = diasl.Cells(1, (i * 2) - 1).Text
    If diacont = "L" Or diacont = "M" Or diacont = "J" Or diacont = "V" Then
     diadd = diadd + 1
    End If
    Next
    Diastrabajables = diadd
End Function
Public Function Diasregulares(diasr As Range, trabajables As Integer)
 Diasregulares = 0
 For i = 1 To diasr.Columns.count
   If diasr.Cells(1, i) <> "" Then
      Diasregulares = Diasregulares + 1
    End If
Next
Diasregulares = Diasregulares / 2
If Diasregulares > trabajables Then
    Diasregulares = trabajables
End If
End Function

Public Function Diasextra(diasr As Range, regulares As Integer)
 Diasextra = 0
 For i = 1 To diasr.Columns.count
   If diasr.Cells(1, i) <> "" Then
      Diasextra = Diasextra + 1
    End If
Next
Diasextra = Diasextra / 2
If Diasextra < regulares Then
    Diasextra = 0
Else
    Diasextra = Int(Diasextra - regulares)
End If
End Function

Public Function HorasLaboradas(diasr As Range, trabajables As Integer)
Diastotaltrabajados = Diasregulares(diasr, trabajables) + Diasextra(diasr, trabajables)

For i = 1 To diasr.Columns.count
    If i Mod 2 = 0 Then
    HorasLaboradas = HorasLaboradas + diasr.Cells(1, i)
    Else
    HorasLaboradas = HorasLaboradas - diasr.Cells(1, i)
    End If
Next

HorasLaboradas = HorasLaboradas - Diastotaltrabajados
End Function

Public Function HorasRegulares(diasr As Range, trabajables As Integer)
    horaslab = HorasLaboradas(diasr, trabajables)
    horastrabaj = trabajables * 8
    horasreg = 0
    If horastrabaj < horaslab Then
        horasreg = horastrabaj
    Else
        horasreg = horaslab
    End If
    HorasRegulares = horasreg
  
End Function

Public Function Horas50(diasr As Range, trabajables As Integer, horas100 As Integer)
    horaslab = HorasLaboradas(diasr, trabajables)
    horasreg = HorasRegulares(diasr, trabajables)
    h50 = horaslab - horasreg - horas100
    
    If h50 < 0 Then
    h50 = 0
    End If
    Horas50 = h50
End Function



Public Function ContarTrabajadores(inIds As String) As Integer
    ContarTrabajadores = 0
    Dim ids() As String
    ids = Split(inIds, ",")
    ContarTrabajadores = UBound(ids) - LBound(ids) + 1
End Function

Public Function BuscarTrabajadores(inIds As String) As String
    BuscarTrabajadores = ""
    Dim ids() As String
    ids = Split(inIds, ",")
   Dim nombre1 As String
   Dim apellido1 As String
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
    nombre1 = WorksheetFunction.VLookup(Val(ids(i)), Worksheets("Trabajadores").Range("B:J"), 8, False)
      apellido1 = WorksheetFunction.VLookup(Val(ids(i)), Worksheets("Trabajadores").Range("B:AD"), 6, False)
        BuscarTrabajadores = BuscarTrabajadores & nombre1 & " " & apellido1 & ","
    
    Next i
End Function

Public Function SepararId(inIds As String, num) As Integer
    SepararId = 0
    Dim ids() As String
    ids = Split(inIds, ",")
    SepararId = ids(num - 1)
End Function


