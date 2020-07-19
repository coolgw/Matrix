Sub aa()

    ' this is used to check generate test case logic
    start_row = Selection.row
    start_col = Selection.column
    
    'Dim tc
    Set tc = New Testcase
    tc.row = start_row
    tc.column = start_col
    tc.generate_case_name
    MsgBox tc.case_name

End Sub

Public Function case_name(row As Integer, column As Integer) As Testcase
    'Dim tc
    Set tc = New Testcase
    tc.row = row
    tc.column = column
    tc.generate_case_name
    Set case_name = tc

End Function


Sub create_case_name_table()

    'Dim start_row, start_col As Long
     start_row = 3
     start_col = 4

     i = 0
     j = 0
    Do While j < 4
     i = 0
     Do While Cells(start_row + i, start_col + j) <> "" ' And i < 5
        Set tc = case_name(start_row + i, start_col + j)
        p1p2 = tc.p1p2
        
        'MsgBox case_name
        Cells(start_row + i, start_col + 5 + j) = tc.case_name
        If p1p2 = "p1" Then
        Cells(start_row + i, start_col + 5 + j).Interior.ColorIndex = 4
        ElseIf p1p2 = "p2" Then
        Cells(start_row + i, start_col + 5 + j).Interior.ColorIndex = 6
        Else
        'Cells(start_row + i, start_col + 5 + j).Interior.ColorIndex = 11
        Cells(start_row + i, start_col + 5 + j) = "/"
        End If
        
        
      'MsgBox "The value of i is : " & case_name & start_row & i
      i = i + 1
     Loop
     j = j + 1
    Loop
     

End Sub


Sub create_setting_table()

End Sub

