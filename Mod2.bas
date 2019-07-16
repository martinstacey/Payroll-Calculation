Attribute VB_Name = "Módulo2"
Function cd(Data_Range As Range) As Variant
    No_Of_Rows_in_Range = Data_Range.Rows.count
    cd = No_Of_Rows_in_Range
End Function


Sub Calendario()

End Sub





Function Buscar2(inrange As Range, Parameter1 As Variant, Col1 As Integer, Parameter2 As Variant, Col2 As Integer, Col As Integer) As Variant

Buscar2 = CVErr(xlErrNA)

For i = 1 To inrange.Rows.count
    If inrange.Cells(i, Col1) = Parameter1 And inrange.Cells(i, Col2) = Parameter2 Then
        Buscar2 = inrange.Cells(i, Col)
    End If
Next

End Function

Public Function Si0blank(infor As String) As String
    If infor = "0" Then
        Si0blank = " "
    Else
        Si0blank = infor
    End If
End Function


Public Function rolidynum(thiscell As Range) As String
        
    myVLookupResult = Application.VLookup(newrolID, arange, 2, False)
        If IsError(myVLookupResult) Then
        
        End If
        
    rolidynum = Application.Caller.Row - 4
    
End Function




Sub CopiarRoles()
Attribute CopiarRoles.VB_ProcData.VB_Invoke_Func = "e\n14"
' CopiarNum Macro
' Acceso directo: CTRL+e
    Sheets("Roles").Select
    Col1 = "B"
    Fil1 = 8
    Med = ":"
    Colm1 = "CT"
    Film1 = Range("j4").Value + Fil1 - 1
    Copyrange = Col1 & Fil1 & Med & Colm1 & Film1
    
    newrolID = Range(Col1 & Fil1).Value
    
    If Range("I3").Value > 0 Then
        Range(Copyrange).Select
        Selection.Copy
        Sheets("Roles Guardados").Select
        Filinit = 5
        ColInit = "B"
        Filfin = Range("C2").Value + Filinit - 1
        ColFin = "CS"
        Dim arange As Range
        Set arange = Range(ColInit & Filinit & Med & ColFin & Filfin)
        arange.Select
        myVLookupResult = Application.VLookup(newrolID, arange, 2, False)
        If IsError(myVLookupResult) Then
            If Range("C2").Value > 0 Then
                NextFil = Filfin
            Else
                NextFil = Filfin + 1
            End If
            Range(ColInit & NextFil).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                        
        Else
            Sheets("Roles").Select
            MsgBox "RolId Repetido Cambiar Id"
        End If
    End If
End Sub

Sub BorrarRoles()
    Sheets("Roles").Select
    Col1 = "I"
    Fil1 = 8
    Med = ":"
    Colm1 = "BV"
    Film1 = Range("j4").Value + Fil1 - 1
    Copyrange = Col1 & Fil1 & Med & Colm1 & Film1
    Delvalsrange = Col1 & Fil1 & Med & Colm1 & Film1
    Range(Delvalsrange).Select
    Range(Copyrange).ClearContents
End Sub

