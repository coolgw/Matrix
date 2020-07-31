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
Public reg As String
Public patch_mode As String
Public lock_mode As String
Public smt_mode As String
Public auto_mode As String
Public base_origin_string As String
Public setting As Dictionary
Public setting_str As String
Public migration_type As String
Public module_name As Dictionary


Public Sub generate_case_name()

    'create_module_name
    init_module_name
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
    'remove carry return
    base_origin_string = Trim(Replace(base_ver, Chr(10), ""))
    base_ver = Replace(base_ver, " ", "_")
    'base_ver = Replace(base_ver, "LTSS", "_LTSS")
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
      lock_mode = "1"
    End If
      
    If InStr(case_mode, "/am") = 5 Then
      auto_mode = "1"
      reg = "media"
      method = "auto"
    End If
      
    If InStr(case_mode, "/ac") = 5 Then
      auto_mode = "1"
      reg = "scc"
      method = "auto"
    End If
    
    ' hpc system role
    hpc_system_role = ""
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
        smt_mode = "1"
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

    'replace recommend for hpc
    If InStr(base_origin_string, "HPC") > 0 Then
        If InStr(addons, "recommended") > 0 Then
            addons = Replace(addons, "recommended", "basesys+desk+dev+hpc+py2+srv+wsm")
        End If
    End If

    ' generate case name and store in case_name
    ' migration_type
    If auto_mode = "1" Then
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
    If smt_mode = "1" Then
        case_name = case_name & "_" & "smt_pattern"
    End If
    ' lockmode
    If lock_mode = "1" Then
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

Public Sub generate_case_setting()

    Set setting = New Dictionary
    ' default setting
    setting_UPGRADE
    setting_DESKTOP


    setting.Add Key:="PATCH", Item:=1

    setting_ROLLBACK_AFTER_MIGRATION
    
    'full or minial update
    If patch_mode = "full" Then
        setting.Add Key:="FULL_UPDATE", Item:=1
    ElseIf patch_mode = "minimal" Then
        setting.Add Key:="MINIMAL_UPDATE", Item:=1
    End If
    
    setting_HDDVERSION
    setting_HDD_1
    setting_ISO
    setting_SLE_PRODUCT
    setting_KEEP_REGISTERED
    
  
    'platform related setting
    'ADDONS need check code, all package?? this seems only related with media upgrade
    If platform = "s390x" Then
          setting.Add Key:="ADDONURL", Item:=addons
    End If
    
    setting_SCC_ADDONS

    'register type
    If reg = "media" Then
        setting.Add Key:="MEDIA_UPGRADE", Item:=1
        setting.Add Key:="ADDONS", Item:="all-packages"
    ElseIf reg = "smt" Then
        setting.Add Key:="SMT_URL", Item:="https://migration-smt.qa.suse.de"
        setting.Add Key:="+SCC_URL", Item:="none"
    ElseIf reg = "rmt" Then
        setting.Add Key:="SMT_URL", Item:="https://openqa-rmt.suse.de"
        setting.Add Key:="+SCC_URL", Item:="none"
    ElseIf reg = "scc" Then
    ElseIf reg = "pscc" Then
    End If

    
    'migration type
    If migration_type = "offline" Then
    ElseIf migration_type = "online" Then
        setting.Add Key:="ONLINE_MIGRATION", Item:=1
        setting.Add Key:="BOOT_HDD_IMAGE", Item:=1
    End If
    'migration method
    If method = "y" Then
        setting.Add Key:="MIGRATION_METHOD", Item:="yast"
    ElseIf method = "zypp" Then
        setting.Add Key:="MIGRATION_METHOD", Item:="zypper"
    ElseIf method = "zdup" Then
        setting.Add Key:="ZDUP", Item:=1
    ElseIf method = "auto" Then
        setting.Add Key:="AUTOYAST", Item:="http://admin.openqa.test/autoyast"
    End If
    
    'lock mode
    If lock_mode = "1" Then
    setting.Add Key:="LOCK_PACKAGE", Item:="zip,sysvinit-tools"
    End If
    
    'PATTERNS
    setting.Add Key:="PATTERNS", Item:=pattern_mode
    
    setting_BOOTFROM
    
    'extra setting for pvm
    setting_pvm
    
    setting_REPO_0
    
    setting_ZDUP
    
End Sub

Public Sub setting_UPGRADE()
    setting.Add Key:="UPGRADE", Item:=1
End Sub
Public Sub setting_DESKTOP()
    ' def desktop is gnome
    setting.Add Key:="DESKTOP", Item:="gnome"
    ' for hpc test group we set textmode for tm and ms systemrole, set textmode for ld
    If hpc_system_role = "ms" Or hpc_system_role = "tm" Then
        setting("DESKTOP") = "textmode"
    End If
    If InStr(case_mode, "textmode") Then
        setting("DESKTOP") = "textmode"
    End If
End Sub
Public Sub setting_HDDVERSION()
           
    Dim arr() As String

    arr = Split(base_origin_string, " ")
    arr = Filter(arr, "SLES", Flase, vbTextCompare)
    arr = Filter(arr, "SLED", Flase, vbTextCompare)
    arr = Filter(arr, "SLE", Flase, vbTextCompare)
    arr = Filter(arr, "LTSS", Flase, vbTextCompare)
    arr = Filter(arr, "SLE", Flase, vbTextCompare)
    arr = Filter(arr, "HPC", Flase, vbTextCompare)
    setting.Add Key:="HDDVERSION", Item:=Trim(UCase(Join(arr, "-")))
    
End Sub

Public Sub setting_HDD_1()
    If InStr(base_origin_string, "HPC") > 0 Then
        'HDD_1 example HDD_1=SLES-12-SP3-%ARCH%-allpatterns-updated.qcow2
        hdd = "SLEHPC-" & setting("HDDVERSION") & "-%ARCH%-" & "GM"
    ElseIf InStr(base_origin_string, "SLED") > 0 Then
        hdd = "SLED-" & setting("HDDVERSION") & "-%ARCH%-" & "GM"
    Else
        hdd = "SLES-" & setting("HDDVERSION") & "-%ARCH%-" & "GM"
    End If
    
    If InStr(base_ver, "15_SP") > 0 And InStr(addons, "dev") > 0 Then
        hdd = hdd & "-" & "SDK"
    ElseIf (InStr(base_ver, "12_SP") > 0 Or InStr(base_ver, "11_SP") > 0) And InStr(addons, "SDK") > 0 Then
        hdd = hdd & "-" & "SDK"
    End If
    
    If hpc_system_role <> "" Then
        If hpc_system_role = "ms" Then
            hdd = hdd & "-" & "SERVER"
        ElseIf hpc_system_role = "ld" Then
            hdd = hdd & "-" & "DEV"
        Else
            hdd = hdd & "-" & "TEXTMODE"
        End If
    End If
    If setting("DESKTOP") = "gnome" Then
        hdd = hdd & "-" & "gnome"
    End If
    If pattern_mode = "all" Then
        hdd = hdd & "-" & "allpatterns"
    End If
    
    setting.Add Key:="HDD_1", Item:=hdd & ".qcow2"
        
End Sub
Public Sub setting_ISO()
    'ISO example ISO_1=SLE-%VERSION%-Packages-%ARCH%-Build%BUILD%-Media1.iso
    '+ISO: SLE-%VERSION%-Full-%ARCH%-Build%BUILD%-Media1.iso
    ' this need radom set?
    If platform <> "s390x" Then
        If (migration_type = "offline" And reg = "media") Or (InStr(case_mode, "/fulldvd") > 0) Then
            setting.Add Key:="+ISO", Item:="SLE-%VERSION%-Full-%ARCH%-Build%BUILD%-Media1.iso"
        Else
            'default no need set since trigger command will take this
            'setting.Add Key:="ISO", Item:="SLE-%VERSION%-Online-%ARCH%-Build%BUILD%-Media1.iso"
        End If
    End If
End Sub
Public Sub setting_SLE_PRODUCT()
    If InStr(base_origin_string, "HPC") > 0 Then
        setting.Add Key:="SLE_PRODUCT", Item:="hpc"
    End If
End Sub

Public Sub setting_SCC_ADDONS()
    'scc addons
    If InStr(base_origin_string, "HPC") > 0 Then
        If InStr(addons, "recommended") > 0 Then
            addons = Replace(addons, "recommended", "basesys+desk+dev+hpc+py2+srv+wsm")
        End If
    End If
    setting.Add Key:="SCC_ADDONS", Item:=Replace(addons, "+", ",")
End Sub

Public Sub setting_ROLLBACK_AFTER_MIGRATION()
    If InStr(case_mode, "rollback") Or InStr(base_origin_string, "HPC") > 0 Then
        setting.Add Key:="ROLLBACK_AFTER_MIGRATION", Item:=1
    End If
End Sub

Public Sub setting_KEEP_REGISTERED()
    If reg <> "media" And migration_type = "offline" Then
        setting.Add Key:="KEEP_REGISTERED", Item:=1
    End If
End Sub

Public Sub setting_BOOTFROM()
    If InStr(case_mode, "/zvm") Or InStr(case_mode, "/pvm") Then

    Else
        setting.Add Key:="BOOTFROM", Item:="d"
    End If
End Sub

Public Sub setting_pvm()
    If InStr(case_mode, "/pvm") Or (platform = "ppc64le" And InStr(base_ver, "15_SP2") > 0) Then  ' 15sp2 ppc64le all pvm case
        If setting.Exists("BOOT_HDD_IMAGE") Then
            setting("BOOT_HDD_IMAGE") = "norm"
        Else
            setting.Add Key:="BOOT_HDD_IMAGE", Item:="norm"
        End If
        setting.Add Key:="SERIALDEV", Item:="hvc0"
        If migration_type = "offline" Then
            setting.Add Key:="YAML_SCHEDULE", Item:="schedule/migration/offline_spvm_Upgrade.yaml"
            setting.Add Key:="MIRROR_HTTP", Item:="http://openqa.suse.de/assets/repo/SLE-15-SP3-Full-ppc64le-Build%BUILD_SLE%-Media1"
        End If
    End If
End Sub

Public Sub setting_REPO_0()
    If platform = "s390x" Or InStr(case_mode, "/pvm") Or (platform = "ppc64le" And InStr(base_ver, "15_SP2") > 0) Then ' fix me for pvm part
        If migration_type = "offline" And reg = "media" Then
            setting.Add Key:="REPO_0", Item:="SLE-%VERSION%-Full-ppc64le-Build%BUILD_SLE%-Media1"
        Else
            setting.Add Key:="REPO_0", Item:="SLE-%VERSION%-Online-ppc64le-Build%BUILD_SLE%-Media1"
        End If
        If InStr(case_mode, "fulldvd") Then  ' flag will override it
            setting("REPO_0") = "SLE-%VERSION%-Full-ppc64le-Build%BUILD_SLE%-Media1"
        End If
    End If
    
    
End Sub

Public Sub setting_ZDUP()
    If method = "zdup" Then
        setting("ZDUP") = "1"
        zdup_repo = ""
        head = "ftp://openqa.suse.de/SLE-%VERSION%-"
        tail = "-%ARCH%-Build%BUILD_SLE%-Media1/"
        'set productor repo
        If InStr(base_ver, "SLED") Then
            prod_repo = "ftp://openqa.suse.de/SLE-%VERSION%-Product-" & "SLED" & "-POOL-%ARCH%-Build%BUILD_SLE%-Media1/"
        ElseIf InStr(base_ver, "SLES") Then
            prod_repo = "ftp://openqa.suse.de/SLE-%VERSION%-Product-" & "SLES" & "-POOL-%ARCH%-Build%BUILD_SLE%-Media1/"
        ElseIf InStr(base_ver, "HPC") Then
            prod_repo = "ftp://openqa.suse.de/SLE-%VERSION%-Product-" & "HPC" & "-POOL-%ARCH%-Build%BUILD_SLE%-Media1/"
        End If
        
        'set iso repo
        If (migration_type = "offline" And reg = "media") Or (InStr(case_mode, "/fulldvd") > 0) Then
            iso_repo = "ftp://openqa.suse.de/SLE-%VERSION%-" & "Online" & "-%ARCH%-Build%BUILD%-Media1/"
        Else
            iso_repo = "ftp://openqa.suse.de/SLE-%VERSION%-" & "Full" & "-%ARCH%-Build%BUILD%-Media1/"
        End If
        
        'set moudle repo
        zdup_repo = zdup_repo & prod_repo & ","
        zdup_repo = zdup_repo & iso_repo & ","
        
       'set module repo
        
        a = Split(addons, "+")
        b = UBound(a)
        For i = 0 To b
           tmp = "ftp://openqa.suse.de/SLE-%VERSION%-Module-" & module_name(a(i)) & "-POOL-%ARCH%-Build%BUILD_SLE%-Media1/"
           zdup_repo = zdup_repo & tmp & ","
        Next
        setting.Add Key:="ZDUPREPOS", Item:=zdup_repo
        
    End If
    
End Sub

Public Sub init_module_name()
    Set module_name = New Dictionary
    module_name.Add Key:="basesys", Item:="Basesystem"
    module_name.Add Key:="srv", Item:="Server-Applications"
    module_name.Add Key:="desk", Item:="Desktop-Applications"
    module_name.Add Key:="dev", Item:="Development-Tools"
    module_name.Add Key:="lp", Item:="Live-Patching"
    module_name.Add Key:="sdk", Item:="SDK"
    module_name.Add Key:="asmm", Item:="Adv-System-Mgmt" '15 not support
    module_name.Add Key:="contm", Item:="Container"
    module_name.Add Key:="lgm", Item:="Legacy"
    module_name.Add Key:="pcm", Item:="Public Cloud"
    module_name.Add Key:="tcm", Item:="Toolchain" '15 not support
    module_name.Add Key:="wsm", Item:="Web-Scripting"
    module_name.Add Key:="phub", Item:="Packagehub"
    module_name.Add Key:="cap", Item:="Cloud-Application-Platform"
    module_name.Add Key:="we", Item:="Product-WE"
    module_name.Add Key:="idu", Item:="IBM-DLPAR-SDK"
    module_name.Add Key:="ids", Item:="IBM-DLPART-Utils"
    module_name.Add Key:="tsm", Item:="Transactional-Server"
    module_name.Add Key:="python2", Item:="Python2"
    
End Sub

