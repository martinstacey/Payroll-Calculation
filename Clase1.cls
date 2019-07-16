VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public inRow, laRow, nRows, nextnRows, tiRow, nextRow As Integer
Public inCol, laCol, nCols, nextnCols As Integer
Public dataRange, insRange, datayinsRange, nextRange, tittleRange, allRange As Range
Public naSheet As String
Private tittles(), values(), nextvalues() As String
Public currentNew, desiredNew, addtoNew As Integer '(PARA ROLES)

Public Sub Instance(IinRow As Integer, ItiRow As Integer, IinCol As Integer, InaSheet As String)
    naSheet = InaSheet
    Sheets(naSheet).Activate
    inRow = IinRow
    tiRow = ItiRow
    inCol = IinCol
    nRows = Application.CountA(Columns(inCol)) - Application.CountA(Range(Cells(1, inCol), Cells(inRow - 1, inCol))) _
                                               - Application.CountIf(Columns(inCol), " ") _
                                               - Application.CountIf(Columns(inCol), "  ") _
                                               - Application.CountIf(Columns(inCol), "   ")

    nCols = Application.CountA(Rows(tiRow)) - Application.CountIf(Rows(inRow), " ") _
                                                - Application.CountIf(Rows(inRow), "  ") _
                                                - Application.CountIf(Rows(inRow), "   ")
    If naSheet = naSheetROL Then nCols = nCols + 31
    If naSheet = naSheetROL Then inRow = inRow + trabaj
    laRow = inRow + nRows - 1
    laCol = inCol + nCols - 1
    nextRow = laRow + 1
    nextnRows = (inRow - 1) - (tiRow + 2) + 1
    nextnCols = nCols
    Set tittleRange = Range(Cells(tiRow, inCol), Cells(tiRow, laCol))
    Set insRange = Range(Cells(tiRow + 2, inCol), Cells(inRow - 1, laCol))
    Set allRange = Range(Columns(inCol), Columns(laCol))
    Set nextRange = Range(Cells(nextRow, inCol), Cells(nextRow + nextnRows - 1, laCol))
    
    If Not nRows = 0 Then
        Set dataRange = Range(Cells(inRow, inCol), Cells(laRow, laCol))
        Set datayinsRange = Union(insRange, dataRange)

        ReDim values(1 To nRows, 1 To nCols)
        Dim r, c As Integer
        For r = 1 To nRows
            For c = 1 To nCols
                values(r, c) = Cells(inRow + r - 1, inCol + c - 1)
            Next c
        Next r
    End If

    ReDim tittles(1 To nCols)
        Dim i As Integer
        For i = 1 To nCols
        tittles(i) = Cells(tiRow, inCol + i - 1)
    Next i

    ReDim nextvalues(1 To nextnRows, 1 To nextnCols)
    Dim nr, nc As Integer
    For nr = 1 To nextnRows
        For nc = 1 To nextnCols
            If Right(tittles(nc), 3) = "(F)" Then
                nextvalues(nr, nc) = Cells(tiRow + 2 + nr - 1, inCol + nc - 1).Formula
            Else
                nextvalues(nr, nc) = Cells(tiRow + 2 + nr - 1, inCol + nc - 1)
            End If
        Next nc
    Next nr
End Sub

Public Function calNewInRowWithInserts() As Integer
    calNewInRowWithInserts = inRow + Application.CountA(Columns(1)) - 1
End Function

Public Sub addColumns(IdesiredNew As Integer)
    desiredNew = IdesiredNew
    currentNew = Application.CountA(Columns(1))
    addtoNew = desiredNew - currentNew
    substracttoNew = -addtoNew
    If addtoNew > 0 Then
    Dim i As Integer
        For i = 1 To addtoNew
            Rows(tiRow + 2).Insert
            Cells(tiRow + 2, 1) = "*"
            Rows(tiRow + 3).Copy
            Rows(tiRow + 2).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            Cells(1, 1).Activate
        Next i
    ElseIf addtoNew < 0 Then
        For i = 1 To substracttoNew
            Rows(tiRow + 2).Delete
        Next i
    End If
End Sub



Public Sub SelectDb()
    Sheets(naSheet).Activate
    dataRange.Select
End Sub
Public Sub SelectInsert()
    Sheets(naSheet).Activate
    insRange.Select
End Sub
Public Sub SelectDins()
    Sheets(naSheet).Activate
    datayinsRange.Select
End Sub
Public Sub SelectTittles()
    Sheets(naSheet).Activate
    tittleRange.Select
End Sub
Public Sub SelectAll()
    Sheets(naSheet).Activate
    allRange.Select
End Sub
Public Sub SelectNext()
    Sheets(naSheet).Activate
    nextRange.Select
End Sub

Public Sub PrintTittles()
    Sheets(naSheet).Activate
    Dim i As Integer
    For i = 1 To nCols
    MsgBox tittles(i)
    Next i
End Sub
Public Sub PrintValues()
    Sheets(naSheet).Activate
    Dim r, c As Integer
    For r = 1 To nRows
        For c = 1 To nCols
            MsgBox values(r, c)
        Next c
    Next r
End Sub
Public Sub PrintNextValues()
    Sheets(naSheet).Activate
    Dim r, c As Integer
    For r = 1 To nextnRows
        For c = 1 To nextnCols
            MsgBox Right(tittles(c), 3) = "(F)"
            'MsgBox nextvalues(r, c)
        Next c
    Next r
End Sub

Public Sub WriteCells()
    Sheets(naSheet).Activate
    Dim r, c As Integer
    For r = 1 To nextnRows
        For c = 1 To nextnCols
            
            Cells(nextRow + r - 1, inCol + c - 1) = nextvalues(r, c)
        Next c
    
    
    
    Next r
End Sub
Public Sub WriteCellsFormulas()
    Sheets(naSheet).Activate
    Dim r, c As Integer
    For r = 1 To nextnRows
        For c = 1 To nextnCols
            Cells(nextRow + r - 1, inCol + c - 1) = nextformulas(r, c)
        Next c
    Next r
End Sub
Public Sub Sort()
    Sheets(naSheet).Activate
    idRange.NumberFormat = "@"
    dataRange.Sort Key1:=idRange, Order1:=xlAscending, Header:=xlNo, DataOption1:=xlSortTextAsNumbers
End Sub

Public Sub SortNum()
    Sheets(naSheet).Activate
    dataRange.Columns(1).NumberFormat = "0"
    dataRange.Sort Key1:=dataRange.Columns(1), Order1:=xlAscending, Header:=xlNo, DataOption1:=xlSortTextAsNumbers
End Sub

Function CheckRepetition(colnum As Integer) As Boolean
    CheckRepetition = False
    Dim nr, r As Integer
    For nr = 1 To nextnRows
       For r = 1 To nRows
            If nextvalues(nr, colnum) = values(r, colnum) Then CheckRepetition = True
        Next r
    Next nr
End Function

Function CheckValueType(colnum As Integer, whatToCheck As String) As Boolean     'isempty,isnotnumber,isnotdate
    CheckValueType = False
    Dim nr, r As Integer
    For nr = 1 To nextnRows
        If whatToCheck = "isempty" And IsEmpty(insRange.Cells(nr, colnum)) Then
            CheckValueType = True
        ElseIf whatToCheck = "isnotnumber" And Not IsNumeric(insRange.Cells(nr, colnum).Value) Then
            CheckValueType = True
        ElseIf whatToCheck = "isnotdate" And Not IsDate(insRange.Cells(nr, colnum).Value) Then
             CheckValueType = True            'No Funciona
        End If
    Next nr
End Function


Function CheckValueValue(colnum As Integer, whatToCheck As String, colnumcomp As Integer) As Boolean    '<,=,>
    CheckValueValue = False
    Dim nr, r As Integer
    For nr = 1 To nextnRows
        If whatToCheck = "<" And insRange.Cells(nr, colnum) < insRange.Cells(nr, colnumcomp) Then CheckValueValue = True
        If whatToCheck = "=" And insRange.Cells(nr, colnum) = insRange.Cells(nr, colnumcomp) Then CheckValueValue = True
        If whatToCheck = ">" And insRange.Cells(nr, colnum) > insRange.Cells(nr, colnumcomp) Then CheckValueValue = True
    Next nr
End Function

Function CheckValueNumber(colnum As Integer, whatToCheck As String, number As Integer) As Boolean    '<,=,>
    CheckValueNumber = False
    Dim nr, r As Integer
    For nr = 1 To nextnRows
        If whatToCheck = "<" And insRange.Cells(nr, colnum) < number Then CheckValueNumber = True
        If whatToCheck = "=" And insRange.Cells(nr, colnum) = number Then CheckValueNumber = True
        If whatToCheck = ">" And insRange.Cells(nr, colnum) > number Then CheckValueNumber = True
    Next nr
End Function




Public Sub RepeatFormatComment(colnum As Integer)
    Sheets(naSheet).Activate
    rowrep = 0
    Dim nr, r As Integer
    For nr = 1 To nextnRows
       For r = 1 To nRows
            If nextvalues(nr, colnum) = values(r, colnum) Then
                rowrep = nr
            End If
        Next r
    Next nr
'    If rowrep > 0 Then
'        With insRange.Cells(rowrep, colnum)
'            .ClearComments
'            .AddComment tittles(colnum) & " Repetido"
'            .Comment.Visible = True
'            .Comment.Shape.ScaleWidth 0.75, msoFalse, msoScaleFromTopLeft
'            .Comment.Shape.ScaleHeight 0.25, msoFalse, msoScaleFromTopLeft
'        End With
'    Else
'        Dim i  As Integer
'        For i = 1 To nextnRows
'            With insRange.Cells(i, colnum)
'                .ClearComments
'            End With
'        Next i
'    End If
    Set oFormat = datayinsRange.FormatConditions.Add(xlExpression, xlFormula, "=CONTAR.SI(" & datayinsRange.Columns(colnum).Address & "," & datayinsRange.Columns(colnum).Cells(1, 1).Address(0, 1) & ")>1")
    oFormat.Interior.Color = RGB(255, 0, 0)
End Sub

Public Sub cellequalFormat(colnum As Integer, condition As String)
    Sheets(naSheet).Activate
    q = Chr(34)
    Set oFormat = dataRange.FormatConditions.Add(xlExpression, xlFormula, "=" & dataRange.Cells(1, colnum).Address(RowAbsolute:=False) & "=" & q & condition & q)
    oFormat.Interior.Color = RGB(255, 200, 200)
End Sub




Function SearchVar(inputval As String, colIndice As Integer, colreturn As Integer) As Variant
    Dim r  As Integer
    SearchVar = "No Encontrado"
    For r = 1 To nRows
        If values(r, colIndice) = inputval Then SearchVar = values(r, colreturn)
    Next r
End Function

Function SearchRow(inputval As String, colIndice As Integer) As Variant
    Dim r  As Integer
    SearchRow = "No Encontrado"
    For r = 1 To nRows
        If values(r, colIndice) = inputval Then SearchRow = r + inRow - 1
    Next r
End Function

