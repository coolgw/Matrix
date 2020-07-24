Sub test()

    ' this is used to check generate test case logic
    start_row = Selection.row
    start_col = Selection.column
    
    'Dim tc
    Set tc = New Testcase
    tc.row = start_row
    tc.column = start_col
    tc.generate_case_name
    tc.generate_case_setting
    ' test case_name
    ' MsgBox tc.case_name
    ' MsgBox tc.setting("PATCH")

    For Each k In tc.setting.Keys
    ' Print key and value
        myline = k & "=" & tc.setting(k) & ""
        setting_str = setting_str & myline & vbCrLf
    
    Next
    ' test setting
    MsgBox setting_str

End Sub

Public Function case_name(row As Integer, column As Integer) As Testcase
    'Dim tc
    Set tc = New Testcase
    tc.row = row
    tc.column = column
    tc.generate_case_name
    tc.generate_case_setting
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
        
        Set tc = Nothing
      'MsgBox "The value of i is : " & case_name & start_row & i
      i = i + 1
     Loop
     j = j + 1
    Loop
     

End Sub

Sub create_setting_table()
    Dim p1_case_setting As String
    Dim p2_case_setting As String
    'Dim start_row, start_col As Long
     start_row = 3
     start_col = 4

     i = 0
     j = 0
    Do While j < 4
     i = 0
     p1_case_setting = ""
     p2_case_setting = ""
     Do While Cells(start_row + i, start_col + j) <> "" ' And i < 5
        Set tc = case_name(start_row + i, start_col + j)
        p1p2 = tc.p1p2
        case_setting_str = tc.case_name & vbCrLf & tc.setting_str & vbCrLf
        'format the content
        case_setting_str = "    - " & tc.case_name & ":" & vbCrLf
        case_setting_str = case_setting_str & "        " & "testsuite: null" & vbCrLf
        case_setting_str = case_setting_str & "        " & "settings:" & vbCrLf
        
        setting_str = "" ' make sure clear the string before rebuild
        For Each k In tc.setting.Keys
        ' Print key and value
            myline = "          " & k & ": '" & tc.setting(k) & "'"
            setting_str = setting_str & myline & vbCrLf
        Next
        case_setting_str = case_setting_str & setting_str
                
        'MsgBox case_name
        Cells(start_row + i, start_col + 5 + j) = case_setting_str
        If p1p2 = "p1" Then
            Cells(start_row + i, start_col + 5 + j).Interior.ColorIndex = 4
            p1_case_setting = p1_case_setting & case_setting_str & vbCrLf
        ElseIf p1p2 = "p2" Then
            Cells(start_row + i, start_col + 5 + j).Interior.ColorIndex = 6
            p2_case_setting = p2_case_setting & case_setting_str & vbCrLf
        Else
            'Cells(start_row + i, start_col + 5 + j).Interior.ColorIndex = 11
            Cells(start_row + i, start_col + 5 + j) = "/"
        End If
        ' set p1 and p2 setting
        Cells(1, start_col + 5 + j) = p1_case_setting
        Cells(2, start_col + 5 + j) = p2_case_setting
        Set tc = Nothing
        
      'MsgBox "The value of i is : " & case_name & start_row & i
      i = i + 1
     Loop
     j = j + 1
    Loop
     
End Sub


Sub create_setting_for_all_sheet()
    Sheets("sles_sled_offline").Activate
    create_setting_table
    Rows("1:100").RowHeight = 20
    Sheets("sles_sled_online").Activate
    create_setting_table
    Rows("1:100").RowHeight = 20
    Sheets("hpc_offline").Activate
    create_setting_table
    Rows("1:100").RowHeight = 20
    Sheets("hpc_online").Activate
    create_setting_table
    Rows("1:100").RowHeight = 20
End Sub
