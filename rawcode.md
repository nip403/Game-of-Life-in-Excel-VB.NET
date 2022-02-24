```
Public Function convert(x)
    convert = Split(Cells(1, x).Address, "$")(1)
End Function
Public Function reset()
    Cells.Select
    Selection = ""
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.Color = vbBlack
End Function
Public Function generate(xMin, xMax, yMin, yMax)
    For x = xMin To xMax
        For y = yMin To yMax
            rn = Int(Rnd * 2)
            Range(convert(x) & y).Select
            ActiveCell.FormulaR1C1 = rn
        Next y
    Next x
End Function
Public Function get_bounds(xMin, xMax, yMin, yMax)
    get_bounds = convert(xMin) & yMin & ":" & convert(xMax) & yMax
End Function
Public Function highlight(xMin, xMax, yMin, yMax)
    For x = xMin To xMax
        For y = yMin To yMax
            If Range(convert(x) & y).Value = 0 Then
                Range(convert(x) & y).Interior.Color = vbBlack
                Range(convert(x) & y).Font.Color = vbBlack
            Else
                Range(convert(x) & y).Interior.Color = vbWhite
                Range(convert(x) & y).Font.Color = vbWhite
            End If
        Next y
    Next x
End Function
Public Function get_grid(xMin, xMax, yMin, yMax)
    cell_range = get_bounds(xMin, xMax, yMin, yMax)
    get_grid = Range(cell_range)
End Function
Public Function next_state(arr, Min, xMax, yMin, yMax)
    Dim next_arr()
    Dim alive_count As Integer
    
    next_arr = arr
    
    For Row = yMin To yMax ' apparently arrays are 1 based?
        For col = xMin To xMax
            alive_count = 0
            
            For nx = -1 To 1
                For ny = -1 To 1
                    If Not nx + ny = 0 Then
                        If (Row + nx > 0 And Row + nx <= dx And col + ny > 0 And col + ny <= dy) Then
                            alive_count = alive_count + 1
                        End If
                    End If
                Next ny
            Next nx
        Next col
    Next Row
End Function
Sub GameOfLife()
    Dim arr(), nextarr()
    Dim xMin As Integer, xMax As Integer, yMin As Integer, yMax As Integer
    Dim endcond As String
    
    ' dimensions of game of life grid
    xMin = 1
    xMax = 5
    yMin = 1
    yMax = 10
    
    Call reset
    Call generate(xMin, xMax, yMin, yMax)
    Call highlight(xMin, xMax, yMin, yMax)
    
    Do While 1
        endcond = InputBox("End?")
        
        If Not endcond = "" Then
            If endcond = "new" Then
                Call generate(xMin, xMax, yMin, yMax)
            Else
                Call reset
                Exit Do
            End If
        End If
        
        arr = get_grid(xMin, xMax, yMin, yMax)
        next_arr = next_state(arr, xMin, xMax, yMin, yMax)
        Range(get_bounds(xMin, xMax, yMin, yMax)) = next_arr
        
        Call highlight(xMin, xMax, yMin, yMax)
    Loop
End Sub
```
