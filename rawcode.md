```
Public Function convert(x)
    convert = Split(Cells(1, x).Address, "$")(1)
End Function
Public Function reset()
    Cells.Select
    Selection = ""
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.Color = vbBlack
    Selection.RowHeight = 15
    Selection.ColumnWidth = 8.43
End Function
Public Function generate(xMin, xMax, yMin, yMax)
    For x = xMin To xMax
        For y = yMin To yMax
            rn = Int(Rnd * 2)
            Range(convert(x) & y).Value = rn
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
Public Function next_state(arr, xMin, xMax, yMin, yMax)
    Dim next_arr()
    Dim alive_count As Integer
    Dim row As Integer, col As Integer
    Dim nx As Integer, ny As Integer
    Dim newx As Integer, newy As Integer
    
    next_arr = arr
    
    For row = 1 To yMax - yMin + 1 ' apparently arrays are 1 based?
        For col = 1 To xMax - xMin + 1
            alive_count = 0
            
            For nx = -1 To 1
                For ny = -1 To 1
                    newx = col + nx
                    newy = row + ny
                    
                    If newx <> col Or newy <> row Then
                        If newx > 0 And newy > 0 And newx <= xMax - xMin + 1 And newy <= yMax - yMin + 1 Then
                            If arr(newy, newx) = 1 Then
                                alive_count = alive_count + 1
                            End If
                        End If
                    End If
                Next ny
            Next nx
            
            If (alive_count = 2 Or alive_count = 3) And arr(row, col) = 1 Then
                next_arr(row, col) = 1
            ElseIf alive_count = 3 And arr(row, col) = 0 Then
                next_arr(row, col) = 1
            Else
                next_arr(row, col) = 0
            End If
        Next col
    Next row
    
    next_state = next_arr
End Function
Public Function resize(xMin, xMax, yMin, yMax)
    Rows(yMin & ":" & yMax).RowHeight = 18.75
    Columns(convert(xMin) & ":" & convert(xMax)).ColumnWidth = 2.86
End Function
Public Function glider(xMin, xMax, yMin, yMax)
    Call reset
    Call resize(xMin, xMax, yMin, yMax)
    
    For x = xMin To xMax
        For y = yMin To yMax
            Range(convert(x) & y).Value = 0
        Next y
    Next x
    
    Range(convert(1 + xMin) & yMin).Value = 1
    Range(convert(2 + xMin) & (1 + yMin)).Value = 1
    Range(convert(xMin) & (2 + yMin) & ":" & convert(2 + xMin) & (2 + yMin)).Value = 1
End Function
Sub GameOfLife()
    Dim arr(), nextarr()
    Dim xMin As Integer, xMax As Integer, yMin As Integer, yMax As Integer
    Dim endcond As String
    
    ' dimensions of game of life grid
    xMin = 1
    xMax = 15
    yMin = 1
    yMax = 15
    
    Call reset
    Call generate(xMin, xMax, yMin, yMax)
    Call highlight(xMin, xMax, yMin, yMax)
    Call resize(xMin, xMax, yMin, yMax)
    
    Do While 1
        endcond = InputBox("End?")
        
        If Not endcond = "" Then
            If endcond = "new" Then
                Call generate(xMin, xMax, yMin, yMax)
            ElseIf endcond = "glider" Then
                Call glider(xMin, xMax, yMin, yMax) ' assuming dimensions are big enough, minimum (3x3)
            Else
                Call reset
                Exit Do
            End If
        End If
        
        Call highlight(xMin, xMax, yMin, yMax)
        
        arr = get_grid(xMin, xMax, yMin, yMax)
        next_arr = next_state(arr, xMin, xMax, yMin, yMax)
        Range(get_bounds(xMin, xMax, yMin, yMax)) = next_arr
    Loop
End Sub
```
