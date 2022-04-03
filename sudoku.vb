Option Explicit

Global abort_puzzle As Boolean
Global search_depth As Integer
Global max_depth As Integer

Sub calculate_hints(hints_values)
    Dim r As Integer, c As Integer
    For r = 1 To 9
        For c = 1 To 9
            hints_values(r, c) = get_cell_hints(r, c)
            DoEvents
        Next c
    Next r
End Sub

Function save_game()
    Dim r As Integer, c As Integer, v As Integer, the_cell As String, board As String
    board = ""
    For r = 1 To 9
        For c = 1 To 9
            v = Int(Sheet1.Cells(r, c).Value)
            If v = 0 Then
                board = board + "."
            Else
                the_cell = Trim(Str(v))
                board = board + the_cell
            End If
            DoEvents
        Next c
    Next r
    save_game = board
End Function

Sub load_game(board As String)
    Dim r As Integer, c As Integer, v As Integer, the_cell As String
    For r = 1 To 9
        For c = 1 To 9
            v = r * 9 + c - 9
            the_cell = Mid(Replace(board, ".", "0"), v, 1)
            If Int("0" & the_cell) = 0 Then
                the_cell = ""
            End If
            Sheet1.Cells(r, c).Value = the_cell
            DoEvents
        Next c
    Next r
End Sub

Function block_number(r As Integer, c As Integer)
    Dim v As Integer, w As Integer
    v = Int((r + 2) / 3)
    w = Int((c + 2) / 3)
    block_number = (v - 1) * 3 + w
End Function

Sub get_block(b As Integer, arr_cells)
    Dim r As Integer, c As Integer, v As Integer, w As Integer, n As Integer, the_cell As String, board As String
    v = (Int((b + 2) / 3) - 1) * 3
    w = (2 + (b - (Int((b + 2) / 3) * 3))) * 3
    For n = 0 To 8
        arr_cells(n) = ""
    Next n
    n = 0
    For r = 1 To 3
        For c = 1 To 3
            arr_cells(n) = "" & Sheet1.Cells(r + v, c + w).Value
            n = n + 1
            DoEvents
        Next c
    Next r
End Sub

Sub get_column(c As Integer, arr_column)
    Dim r As Integer, the_cell As String
    For r = 1 To 9
        the_cell = "" & Sheet1.Cells(r, c).Value
        arr_column(r - 1) = the_cell
        DoEvents
    Next r
End Sub

Sub get_row(r As Integer, arr_row)
    Dim c As Integer, the_cell As String
    For c = 1 To 9
        the_cell = "" & Sheet1.Cells(r, c).Value
        arr_row(c - 1) = the_cell
        DoEvents
    Next c
End Sub

Function units_solved(arr_row)
    Dim v As Integer, solved As Boolean
    solved = True
    For v = 1 To 9
        If arr_row(v - 1) = "" Then
            solved = False
        End If
        DoEvents
    Next v
    units_solved = solved
End Function

Function board_solved()
    Dim cell_units As String, n As Integer, solved As Boolean, arr_row(9) As String
    solved = True
    For n = 1 To 9
        Call get_row(n, arr_row)
        solved = solved And units_solved(arr_row)
        DoEvents
    Next n
    For n = 1 To 9
        Call get_column(n, arr_row)
        solved = solved And units_solved(arr_row)
        DoEvents
    Next n
    For n = 1 To 9
        Call get_block(n, arr_row)
        solved = solved And units_solved(arr_row)
        DoEvents
    Next n
    board_solved = solved
End Function

Function board_sum()
    Dim r As Integer, c As Integer, v As Integer
    v = 0
    For r = 1 To 9
        For c = 1 To 9
            v = v + Int("0" & Sheet1.Cells(r, c).Value)
            DoEvents
        Next c
    Next r
    board_sum = v
End Function

Function count_arr_cells(arr_cells)
    Dim r As Integer, v As Integer, n As Integer
    n = 0
    For r = 1 To 9
        v = Int("0" & arr_cells(r - 1))
        If v > 0 Then
            n = n + 1
        End If
        DoEvents
    Next r
    count_arr_cells = n
End Function

Function hints_available(hints_values)
    Dim r As Integer, c As Integer, v As Integer
    v = 0
    For r = 1 To 9
        For c = 1 To 9
            If Len(hints_values(r, c)) > 0 Then
                v = 1
                r = 10
                Exit For
            End If
            DoEvents
        Next c
    Next r
    hints_available = (v > 0)
End Function

Function get_cell_hints(r As Integer, c As Integer)
    Dim n As Integer, cell_units As String, filled_digits As String, digit As String, arr_row(9) As String
    get_cell_hints = ""
    filled_digits = ""
    n = Int(Sheet1.Cells(r, c).Value)
    If n = 0 Then
        n = block_number(r, c)
        Call get_row(r, arr_row)
        filled_digits = Join(arr_row, "")
        Call get_column(c, arr_row)
        filled_digits = filled_digits + Join(arr_row, "")
        Call get_block(n, arr_row)
        filled_digits = filled_digits + Join(arr_row, "")
        For n = 1 To 9
            digit = Trim(Str(n))
            If InStr(filled_digits, digit) = 0 Then
                get_cell_hints = get_cell_hints + digit
            End If
            DoEvents
        Next
    End If
End Function

Sub optimize_hints(arr_hints, hints_values)
    Dim r As Integer, c As Integer, i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
    Dim digits(9) As String, hint As String, s As String, t As String, digit_sum(9) As Integer
    
    For c = 1 To 9
        For n = 1 To 9
            m = 0
            s = Trim(Str(n))
            For r = 1 To 9  'any single digit hint within column c ?
                DoEvents
                If InStr(hints_values(r, c), s) > 0 Then
                    m = m + 1
                End If
            Next r
            If m = 1 Then
                For r = 1 To 9
                    DoEvents
                    If InStr(hints_values(r, c), s) > 0 Then
                        hints_values(r, c) = s
                    End If
                Next r
            End If
        Next n
    Next c
    
    For r = 1 To 9
        For n = 1 To 9
            m = 0
            s = Trim(Str(n))
            For c = 1 To 9  'any single digit hint within row r ?
                DoEvents
                If InStr(hints_values(r, c), s) > 0 Then
                    m = m + 1
                End If
            Next c
            If m = 1 Then
                For c = 1 To 9
                    DoEvents
                    If InStr(hints_values(r, c), s) > 0 Then
                        hints_values(r, c) = s
                    End If
                Next c
            End If
        Next n
    Next r
       
    For i = 1 To 8
        DoEvents
        s = arr_hints(i)
        If Len(s) = 2 Then
            For j = i + 1 To 9
                DoEvents
                t = arr_hints(j)
                If (Len(t) = 2) And (s = t) Then 'find hidden twins
                    For k = 1 To 9
                        DoEvents
                        If (k <> i) And (k <> j) Then
                            arr_hints(k) = Replace(arr_hints(k), Left(s, 1), "")
                            arr_hints(k) = Replace(arr_hints(k), Right(s, 1), "")
                        End If
                    Next k
                    i = 10
                    j = 10
                    Exit For
                End If
            Next j
        End If
    Next i
       
    'any exactly twins of same 2 digits hints ?
    For r = 1 To 9
        digit_sum(r) = 0
        DoEvents
    Next r
    For m = 1 To 9
        s = Trim(Str(m))
        For n = 1 To 9
            DoEvents
            hint = arr_hints(n - 1)
            If InStr(hint, s) > 0 Then
                digit_sum(m) = digit_sum(m) + 1
            End If
        Next n
    Next m
    j = 0
    For r = 1 To 9
        If digit_sum(r) = 2 Then
            j = j + 1
        End If
        DoEvents
    Next r
    If j = 2 Then
        j = 0
        For r = 1 To 9
            If digit_sum(r) = 2 Then
                digits(j) = Trim(Str(r))
                j = j + 1
            End If
            DoEvents
        Next r
        j = 0
        For r = 1 To 9
            hint = arr_hints(r - 1)
            If (InStr(hint, digits(0)) > 0) And (InStr(hint, digits(1)) > 0) Then
                j = j + 1
            End If
            DoEvents
        Next r
        If j = 2 Then
            For r = 1 To 9
                hint = arr_hints(r - 1)
                If (InStr(hint, digits(0)) > 0) And (InStr(hint, digits(1)) > 0) Then
                    hint = digits(0) + digits(1)
                Else
                    hint = Replace(hint, digits(0), "")
                    hint = Replace(hint, digits(1), "")
                End If
                arr_hints(r - 1) = hint
                DoEvents
            Next r
        End If
    End If
    
    'any exactly 3 pairs of same 3 digits hints ?
    j = 0
    For r = 1 To 9
        If (digit_sum(r) = 2) Or (digit_sum(r) = 3) Then    'bugs : need new logic blocks for '3'
            j = j + 1
        End If
        DoEvents
    Next r
    If j = 3 Then
        j = 0
        For r = 1 To 9
            If (digit_sum(r) = 2) Or (digit_sum(r) = 3) Then
                digits(j) = Trim(Str(r))
                j = j + 1
            End If
            DoEvents
        Next r
        If j = 3 Then
            j = 0
            For r = 1 To 9
                hint = arr_hints(r - 1)
                If (InStr(hint, digits(0)) > 0) Or (InStr(hint, digits(1)) > 0) Or (InStr(hint, digits(2)) > 0) Then
                    j = j + 1
                End If
                DoEvents
            Next r
            If j = 3 Then
                For r = 1 To 9
                    hint = arr_hints(r - 1)
                    If (InStr(hint, digits(0)) > 0) Or (InStr(hint, digits(1)) > 0) Or (InStr(hint, digits(2)) > 0) Then
                        s = ""
                        If (InStr(hint, digits(0)) > 0) Then
                            s = s + digits(0)
                        End If
                        If (InStr(hint, digits(1)) > 0) Then
                            s = s + digits(1)
                        End If
                        If (InStr(hint, digits(2)) > 0) Then
                            s = s + digits(2)
                        End If
                        hint = s
                    Else
                        hint = Replace(hint, digits(0), "")
                        hint = Replace(hint, digits(1), "")
                        hint = Replace(hint, digits(2), "")
                    End If
                    arr_hints(r - 1) = hint
                    DoEvents
                Next r
            End If
        End If
    End If
    
    'segment of 3 units check
    For r = 1 To 9
        DoEvents
        For i = 1 To 3
            DoEvents
            s = ""
            k = 1
            For j = 1 To 3
                DoEvents
                c = j + i * 3 - 3
                s = s + hints_values(r, c)
                If hints_values(r, c) = "" Then
                    k = 0
                End If
            Next j
            If k = 0 Then
                s = ""
            End If
            k = 0
            For j = 1 To 9
                DoEvents
                If InStr(s, Trim(Str(j))) > 0 Then
                    k = k + 1
                    digits(k) = j
                End If
            Next j
            If k = 3 Then
                For c = 1 To 9
                    DoEvents
                    j = Int((c + 2) / 3)
                    If j <> i Then
                        For k = 1 To 3
                            hints_values(r, c) = Replace(hints_values(r, c), digits(k), "")
                            DoEvents
                        Next k
                    End If
                Next c
                For j = 1 To 9
                    DoEvents
                    m = Int((j + 2) / 3) + Int((r + 2) / 3) * 3 - 3
                    c = i * 3 + j - Int((j + 2) / 3) * 3
                    If m <> r Then
                        For k = 1 To 3
                            DoEvents
                            hints_values(m, c) = Replace(hints_values(m, c), digits(k), "")
                        Next k
                    End If
                Next j
            End If
        Next i
    Next r
    
    'need more methods....
End Sub

Function valid_hints(hints_values)
    Dim r As Integer, c As Integer, i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
    Dim arr_hints(9) As String, hint As String, s As String, hint_ok As Boolean
    hint_ok = True
    For m = 1 To 3
        For n = 1 To 3
            s = ""
            For r = 1 To 3
                For c = 1 To 3
                    DoEvents
                    j = r + m * 3 - 3
                    k = c + n * 3 - 3
                    s = s & hints_values(j, k) & Sheet1.Cells(j, k).Value
                Next c
                DoEvents
            Next r
            For r = 1 To 9
                If InStr(s, Trim(Str(r))) = 0 Then
                    hint_ok = False
                    Exit For
                End If
            Next r
            If Not hint_ok Then
                Exit For
            End If
        Next n
        If Not hint_ok Then
            Exit For
        End If
    Next m
    valid_hints = hint_ok
End Function

Sub improve_hints(hints_values)
    Dim r As Integer, c As Integer, i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
    Dim arr_hints(9) As String, hint As String, hint_list As Object, digit_sum(9) As Integer, s As String
    For m = 1 To 3
        For n = 1 To 3
            DoEvents
            For i = 1 To 9
                DoEvents
                digit_sum(i) = 0
                s = Trim(Str(i))
                For r = 1 To 3
                    For c = 1 To 3
                        DoEvents
                        j = r + m * 3 - 3
                        k = c + n * 3 - 3
                        If InStr(hints_values(j, k), s) > 0 Then
                            digit_sum(i) = digit_sum(i) + 1
                        End If
                    Next c
                    DoEvents
                Next r
                If digit_sum(i) = 1 Then
                    For r = 1 To 3
                        For c = 1 To 3
                            DoEvents
                            j = r + m * 3 - 3
                            k = c + n * 3 - 3
                            If InStr(hints_values(j, k), s) > 0 Then
                                hints_values(j, k) = s
                            End If
                        Next c
                        DoEvents
                    Next r
                    DoEvents
                End If
            Next
        Next n
    Next m
    Call apply_hints(hints_values)
    
    For r = 1 To 9
        For c = 1 To 9
            arr_hints(c - 1) = hints_values(r, c)
            DoEvents
        Next c
        If Len(Replace(Join(arr_hints, ","), ",", "")) > 0 Then
            Call optimize_hints(arr_hints, hints_values)
            For c = 1 To 9
                hints_values(r, c) = arr_hints(c - 1)
                DoEvents
            Next c
        End If
    Next
    Call apply_hints(hints_values)
    
    For c = 1 To 9
        For r = 1 To 9
            arr_hints(r - 1) = hints_values(r, c)
            DoEvents
        Next r
        If Len(Replace(Join(arr_hints, ","), ",", "")) > 0 Then
            Call optimize_hints(arr_hints, hints_values)
            For r = 1 To 9
                hints_values(r, c) = arr_hints(r - 1)
                DoEvents
            Next r
        End If
    Next c
    Call apply_hints(hints_values)
    
    For m = 1 To 3
        For n = 1 To 3
            DoEvents
            k = 0
            For r = 1 To 3
                For c = 1 To 3
                    arr_hints(k) = hints_values(r + m * 3 - 3, c + n * 3 - 3)
                    k = k + 1
                    DoEvents
                Next c
                DoEvents
            Next r
            If Len(Replace(Join(arr_hints, ","), ",", "")) > 0 Then
                Call optimize_hints(arr_hints, hints_values)
                k = 0
                For r = 1 To 3
                    For c = 1 To 3
                        hints_values(r + m * 3 - 3, c + n * 3 - 3) = arr_hints(k)
                        k = k + 1
                        DoEvents
                    Next c
                    DoEvents
                Next r
            End If
        Next n
    Next m
    Call apply_hints(hints_values)
End Sub

Sub apply_hints(hints_values)
    Dim r As Integer, c As Integer, i As Integer, j As Integer, m As Integer, n As Integer, err As Integer
    Dim the_cell As String, hints As String, arr_row(9) As String
    m = board_sum()
    If m = 405 Then
        Exit Sub
    End If
    For r = 1 To 9
        For c = 1 To 9
            hints = hints_values(r, c)
            n = Len(hints)
            If n = 1 Then
                err = 0
                Call get_row(r, arr_row)
                the_cell = Join(arr_row, ",")
                If InStr(the_cell, hints) > 0 Then
                    err = 1
                End If
                Call get_column(c, arr_row)
                the_cell = Join(arr_row, ",")
                If InStr(the_cell, hints) > 0 Then
                    err = 1
                End If
                n = block_number(r, c)
                Call get_block(n, arr_row)
                the_cell = Join(arr_row, ",")
                If InStr(the_cell, hints) > 0 Then
                    err = 1
                End If
                If err = 0 Then
                    Sheet1.Cells(r, c).Select
                    Sheet1.Cells(r, c).Value = hints
                    m = m + Int("0" & hints)
                    For m = 1 To 9
                        DoEvents
                        hints_values(r, m) = Replace(hints_values(r, m), hints, "")
                    Next m
                    For m = 1 To 9
                        DoEvents
                        hints_values(m, c) = Replace(hints_values(m, c), hints, "")
                    Next m
                    For m = 1 To 9
                        DoEvents
                        i = Int((r + 2) / 3) * 3 + Int((m + 2) / 3) - 3
                        j = Int((c + 2) / 3) * 3 + m - Int((m + 2) / 3) * 3
                        hints_values(i, j) = Replace(hints_values(i, j), hints, "")
                    Next m
                    
                    hints_values(r, c) = ""
                End If
                If m = 405 Then
                    Exit For
                End If
            End If
            If m = 405 Then
                Exit For
            End If
            DoEvents
        Next c
    Next r
End Sub

Function check_cell_units(arr_row)
    Dim cell_units As String, i As Integer, j As Integer, m As Integer
    m = 0
    For i = 0 To 8
        If arr_row(i) <> "" Then
            For j = (i + 1) To 8
               If arr_row(i) = arr_row(j) Then
                    i = 9
                    m = 1
                    Exit For
               End If
            Next j
        End If
    Next i
    check_cell_units = m
End Function

Function check_board()
    Dim k As Integer, m As Integer, n As Integer, cell_units As String, errmsg As String, arr_row(9) As String
    k = 0
    errmsg = ""
    For n = 1 To 9
        Call get_row(n, arr_row)
        m = check_cell_units(arr_row)
        If m > 0 Then
            errmsg = errmsg + "row " + Str(n) + vbCrLf
            k = 1
        End If
        DoEvents
    Next n
    For n = 1 To 9
        Call get_column(n, arr_row)
        m = check_cell_units(arr_row)
        If m > 0 Then
            errmsg = errmsg + "column " + Str(n) + vbCrLf
            k = 1
        End If
        DoEvents
    Next n
    For n = 1 To 9
        Call get_block(n, arr_row)
        m = check_cell_units(arr_row)
        If m > 0 Then
            errmsg = errmsg + "block " + Str(n) + vbCrLf
            k = 1
        End If
        DoEvents
    Next n
    check_board = True
    If k > 0 Then
        check_board = False
    End If
End Function

Function solve_puzzle()
    Dim cnt As Integer, m As Integer, n As Integer, units_sum As Integer, prev_units_sum As Integer
    Dim solved As Boolean, hint_avail As Boolean, r As Integer, c As Integer, hints_values(9, 9) As String
    cnt = 20
    m = 0
    n = 0
    solved = False
    prev_units_sum = 0
    units_sum = prev_units_sum
    If units_sum = 405 Then
        solve_puzzle = True
        Exit Function
    End If
    Do While (solved = False) And (n < cnt) And (units_sum < 405) And (m = 0) And (Not abort_puzzle)
        Call calculate_hints(hints_values)
        If Not valid_hints(hints_values) Then
            Exit Do
        End If
        Call improve_hints(hints_values)
        DoEvents
        n = n + 1
        units_sum = board_sum()
        hint_avail = hints_available(hints_values)
        If (units_sum <> 405) And (hint_avail = False) Then
            solve_puzzle = False
            Exit Function
        End If
        If (prev_units_sum = units_sum) Then
            m = m + 1
        Else
            m = 0
        End If
        prev_units_sum = units_sum
        solved = board_solved()
    Loop
    solved = board_solved()
    solve_puzzle = solved
End Function

Function travel_cells(game_board As String)
    Dim r As Integer, c As Integer, i As Integer, j As Integer, k As Integer, m As Integer
    Dim hints As String, digit As String, game_board2 As String, solved As Boolean, hints_values(9, 9) As String
    If abort_puzzle Or (search_depth > max_depth) Then
        travel_cells = False
        Exit Function
    End If
    Call calculate_hints(hints_values)
    If Not valid_hints(hints_values) Then
        travel_cells = False
        Exit Function
    End If
    m = 1
    For r = 1 To 9
        For c = 1 To 9
            If (Sheet1.Cells(r, c).Value = "") And (Len(hints_values(r, c)) = 0) Then
                m = 0
                Exit For
            End If
        Next c
        If m = 1 Then
            Exit For
        End If
        DoEvents
    Next r
    If m = 0 Then
        travel_cells = False
        Exit Function
    End If
    m = 0
    For r = 1 To 9
        For c = 1 To 9
            If (Sheet1.Cells(r, c).Value = "") And (Len(hints_values(r, c)) > 0) Then
                m = 1
                Exit For
            End If
        Next c
        If m = 1 Then
            Exit For
        End If
        DoEvents
    Next r
    If m = 0 Then
        travel_cells = False
        Exit Function
    End If
    hints = hints_values(r, c)
    For m = 1 To Len(hints)
        DoEvents
        digit = Mid(hints, m, 1)
        Sheet1.Cells(r, c).Value = digit
        If check_board() Then
            hints_values(r, c) = ""
            solved = solve_puzzle()
            If solved Then
                travel_cells = True
                Exit Function
            End If
            game_board2 = save_game()
            Call calculate_hints(hints_values)
            search_depth = search_depth + 1
            solved = travel_cells(game_board2)
            search_depth = search_depth - 1
            If solved Then
                travel_cells = True
                Exit Function
            End If
        Else
            Sheet1.Cells(r, c).Value = ""
            hints_values(r, c) = hints
        End If
        Call load_game(game_board)
    Next m
    travel_cells = solved
    If Not solved Then
        Call load_game(game_board)
    End If
End Function

Sub solve_using_hints()
    Dim solved As Boolean, msg As String
    abort_puzzle = False
    search_depth = 1
    max_depth = Sheet1.Cells(12, 9).Value
    If Not check_board() Then
        MsgBox "The puzzle is incorrect", vbExclamation, "duplicated digits found"
        Exit Sub
    End If
    Sheet1.lblMsg.Visible = False
    solved = solve_puzzle()
    If Not solved Then
        solved = travel_cells(save_game())
    End If
    If solved Then
        If Int("0" & Sheet1.Cells(16, 9).Value) > 0 Then
            Sheet1.Cells(Sheet1.Cells(16, 9).Value, 8).Value = "'" & save_game()
        End If
		msg = "Sudoku Solver :" + vbCrLf + "Puzzle has been solved !"
        Sheet1.lblMsg.Caption = msg
        Sheet1.lblMsg.Visible = True
    Else
        If MsgBox("Solve it using advanced Python solution ?", vbQuestion + vbYesNo, "Python Solver") = vbYes Then
            Call solve_using_python
            Sheet1.Cells(Sheet1.Cells(16, 9).Value, 8).Value = "'" & save_game()
        Else
            msg = "Sudoku Solver :" + vbCrLf + "Puzzle is not solved !"
            msg = msg + vbCrLf + "Try to increase the max depth" + vbCrLf + "using spin button at I13."
            Sheet1.lblMsg.Caption = msg
            Sheet1.lblMsg.Visible = True
        End If
    End If
End Sub

Sub solve_using_python()
    Dim puzzle As String
    puzzle = "" & Sheet1.Cells(19, 3).Value
    puzzle = python_solver(puzzle)
    Call load_game(puzzle)
End Sub

Function python_solver(puzzle)
    Dim oShell As Object, oCmd As String
    Dim oExec As Object, oOutput As Object
    Dim s As String, sLine As String
    Set oShell = CreateObject("WScript.Shell")
    oCmd = "python sudoku.py" & " " & puzzle
    Set oExec = oShell.Exec(oCmd)
    Set oOutput = oExec.StdOut
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbNewLine
    Wend
    Set oOutput = Nothing: Set oExec = Nothing
    Set oShell = Nothing
    python_solver = sLine
End Function
