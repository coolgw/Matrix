Public row As Integer
Public column As Integer
Public base_ver   As String
Public addons As String
Public pattern_mode As String
Public case_mode As String
Public platform As String
Public case_name As String
Public method As String
Public hpc_system_role As String
Public p1p2 As String
Public Sub generate_case_name()

    'get online or offline
    'MsgBox ActiveSheet.Name
    
    If ActiveSheet.Name Like "*offline*" Then
        migration_type = "offline"
    ElseIf ActiveSheet.Name Like "*online*" Then
        migration_type = "online"
    End If
        
    'get base_ver
    base_ver = Cells(row, 1)
    'MsgBox base_ver
    i = 1
    Do While base_ver = "" And i < 20
      base_ver = Cells(row - i, 1)
      'MsgBox "The value of i is : " & base_ver
      i = i + 1
    Loop
    base_ver = Replace(base_ver, " ", "")
    base_ver = Replace(base_ver, "LTSS", "_LTSS")
    'MsgBox base_ver
    'get moudles
    addons = Cells(row, 2)
    'MsgBox addons
    i = 1
    Do While addons = "" And i < 20
      addons = Cells(row - i, 2)
      'MsgBox "The value of i is : " & addons & row & i
      i = i + 1
    Loop

    'get pattern
    pattern_mode = Cells(row, 3)
    i = 1
    Do While pattern_mode = "" And i < 20
      pattern_mode = Cells(row - i, 3)
      'MsgBox "The value of i is : " & pattern_mode & row & i
      i = i + 1
    Loop

    'get scc mode
    case_mode = Cells(row, column)
    If InStr(case_mode, "p1/c") = 1 Then
     reg = "pscc"
    ElseIf InStr(case_mode, "p2/c") = 1 Then
     reg = "scc"
    ElseIf InStr(case_mode, "p2/s") = 1 Then
      reg = "smt"
    ElseIf InStr(case_mode, "p2/r") = 1 Then
      reg = "rmt"
    End If
    'get platform
    platform = Cells(1, column)
    
    If InStr(case_mode, "p1") = 1 Then
     p1p2 = "p1"
    ElseIf InStr(case_mode, "p2") = 1 Then
     p1p2 = "p2"
    End If
    
    'trim string get from cell
    base_ver = Replace(base_ver, Chr(10), "")
    base_ver = Trim(base_ver)
    addons = Replace(addons, Chr(10), "")
    addons = Trim(addons)
    pattern_mode = Replace(pattern_mode, Chr(10), "")
    pattern_mode = Trim(pattern_mode)
    case_mode = Replace(case_mode, Chr(10), "")
    case_mode = Trim(case_mode)
    platform = Replace(platform, Chr(10), "")
    platform = Trim(platform)
    
    If InStr(case_mode, "/m") = 5 Then
      reg = "media"
    End If
      
    If InStr(case_mode, "/lock") = 7 Then
      lock_mode = 1
    End If
      
    If InStr(case_mode, "/am") = 5 Then
      auto_mode = 1
      reg = "media"
      method = "auto"
    End If
      
    If InStr(case_mode, "/ac") = 5 Then
      auto_mode = 1
      reg = "scc"
      method = "auto"
    End If
    
    ' hpc system role
    If InStr(case_mode, "/ld") > 0 Then
        hpc_system_role = "ld"
    End If
    If InStr(case_mode, "/ms") > 0 Then
        hpc_system_role = "ms"
    End If
    If InStr(case_mode, "/tm") > 0 Then
        hpc_system_role = "tm"
    End If
    
    
    'get method, normally offline case is yast
    If InStr(case_mode, "/y") = 5 Then
      method = "y"
    End If
    If InStr(case_mode, "/z") = 5 Then
      method = "zypp"
    End If
    If InStr(case_mode, "/d") = 5 Then
      method = "zdup"
    End If
    'If InStr(case_mode, "/l") = 7 Then
    '  method = method & "_" & "lock"
    'End If
    
   'default patch_mode is full
    patch_mode = "full"
    'default smt_mode is 0
    smt_mode = 0
    If addons = "Minimal" Then
        patch_mode = "minimal"
        addons = "Base"
    End If
    
    If addons = "SMT Pattern" Then
        smt_mode = 1
        addons = "Base"
    End If
    
    ' special handel part for addons, some addons not support on specific platform
    ' aarch64 need remove asmm, contm, ppc64le need remove we
    Dim arr() As String
    Dim filter_arr() As String
    If platform = "ppc64le" And InStr(addons, "WE") > 0 Then
        arr = Split(addons, "+")
        filter_arr = Filter(arr, "we", Flase, vbTextCompare)
        addons = Join(filter_arr, "+")
    End If
    If platform = "aarch64" And InStr(addons, "asmm") > 0 Then
        arr = Split(addons, "+")
        filter_arr = Filter(arr, "asmm", Flase, vbTextCompare)
        addons = Join(filter_arr, "+")
    End If
    If platform = "aarch64" And InStr(addons, "contm") > 0 Then
        arr = Split(addons, "+")
        filter_arr = Filter(arr, "contm", Flase, vbTextCompare)
        addons = Join(filter_arr, "+")
    End If

    ' generate case name and store in case_name
    ' migration_type
    If auto_mode = 1 Then
        case_name = "autoupgrade"
    Else
        case_name = migration_type
    End If
    ' base_ver + reg + addons + pattern_mode + patch_mode
    case_name = case_name & "_" & base_ver & "_" & reg & "_" & addons & "_" & pattern_mode & "_" & patch_mode
    ' method
    If method <> "" Then
        case_name = case_name & "_" & method
    End If
    ' smt_pattern
    If smt_mode = 1 Then
        case_name = case_name & "_" & "smt_pattern"
    End If
    ' lockmode
    If lock_mode = 1 Then
        case_name = case_name & "_" & "lock"
    End If
    ' hpc ld/ms/tm mode
    If hpc_system_role <> "" Then
        case_name = case_name & "_" & hpc_system_role
    End If
    
    ' no case
    If case_mode = "-" Or case_mode = "/" Or case_mode = "" Then
        case_name = ""
    End If


End Sub

