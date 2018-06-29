VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_hack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=====================================SETTING=====================================
Private Const PreAllocSize As Long = &H10000 * 8
'=======================================API=======================================
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcessAPI Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemoryAPI Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemoryAPI Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandleAPI Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Boolean
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Boolean
Private Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'=====================================Consts======================================
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_VM_OPERATION = &H8
Private Const MEM_COMMIT = &H1000
Private Const MEM_RELEASE = &H8000
Private Const PAGE_READWRITE = &H4
'======================================Types======================================
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
'==================================VarsForClass===================================
Private PreAllocAddress As Long
Private AllocedSize As Long
'======================================Vars=======================================
Public Inited As Boolean
Public hWnd, handle, Pid As Long

Public Function OpenProcess(Optional ByVal lpPID As Long = -1) As Long
    If lpPID = 0 And Pid = 0 Then Exit Function
    If lpPID > 0 And Pid = 0 Then Pid = lpPID
    handle = OpenProcessAPI(PROCESS_ALL_ACCESS, False, Pid)
    OpenProcess = handle
    If handle > 0 Then Inited = True
End Function

Public Function OpenProcessByProcessName(ByVal lpName As String) As Long
    Dim PE32        As PROCESSENTRY32
    Dim hSnapshot   As Long
    Pid = 0
    hSnapshot = CreateToolhelp32Snapshot(2, 0&) 'TH32CS_SNAPPROCESS = 2
    PE32.dwSize = Len(PE32)
    Process32First hSnapshot, PE32
    While Pid = 0 And CBool(Process32Next(hSnapshot, PE32))
        If Right$(LCase$(Left$(PE32.szexeFile, InStr(1, PE32.szexeFile, Chr(0)) - 1)), Len(lpName)) = LCase$(lpName) Then
            Pid = PE32.th32ProcessID
        End If
    Wend
    CloseHandleAPI hSnapshot
    OpenProcessByProcessName = OpenProcess
End Function

Public Function OpenProcessByWindow(ByVal lpWindowName As String, Optional ByVal lpClassName As String = vbNullString) As Long
    hWnd = FindWindow(lpClassName, lpWindowName)
    GetWindowThreadProcessId hWnd, Pid
    OpenProcessByWindow = OpenProcess
End Function

Public Function CloseHandle() As Long
    If Not Inited Then Exit Function
    CloseHandle = CloseHandleAPI(handle)
    handle = 0
    hWnd = 0
    Pid = 0
    Inited = False
End Function

Public Function CloseGame() As Long
    If Not Inited Then Exit Function
    CloseGame = TerminateProcess(handle, 0&)
    CloseHandle
    handle = 0
    Pid = 0
    hWnd = 0
End Function

Public Function ByPassHs(Optional ByVal lpImagePath As String = "C:\WINDOWS\system32\svchost.exe") As Long
    Static BeUsed As Boolean
    If BeUsed = False Then
        Dim hProcess As Long
        hProcess = OpenProcessAPI(PROCESS_ALL_ACCESS, 0, GetCurrentProcessId)
        If hProcess = 0 Then Exit Function
        Dim sLenth As Long
        Dim BaseAddress As Long

If InStr(Environ("ProgramFiles"), "x86") <= 0 Then

        sLenth = LenB(lpImagePath) + 1 + &H1A
       
        BaseAddress = VirtualAllocEx(hProcess, ByVal 0&, ByVal sLenth, MEM_COMMIT, PAGE_READWRITE)
        If BaseAddress = 0 Then Exit Function
        
       'Fuck32Bit
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H0, "64 8B 05 30 00 00 00"                      '//=> | mov eax,fs:[00000030]
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H7, "8B 40 10"                                  '//+7 | mov eax,[eax+10]
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &HA, "83 C0 3C"                                  '//+A | add eax,3C
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &HD, "8B 5C 24 04"                               '//+D | mov ebx,[esp+04]
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H11, "89 18"                                    '//+11 | mov [eax],ebx
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H13, "89 44 24 04"                              '//+13 | mov [esp+04],eax
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H17, "C2 10 00"                                 '//+17 | ret 0010
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + &H1A, ByVal StrPtr(lpImagePath), sLenth, False  '//+1A | inc esp ¼g¤JImagePathName
        
        CloseHandleAPI hProcess
        CallWindowProc BaseAddress, BaseAddress + &H1A, 0, 0, 0
        
        Else
        
        sLenth = LenB(lpImagePath) + 1 + &H1F
       
        BaseAddress = VirtualAllocEx(hProcess, ByVal 0&, ByVal sLenth, MEM_COMMIT, PAGE_READWRITE)
        If BaseAddress = 0 Then Exit Function
        
       'Fuck64Bit
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H0, "64 8B 05 30 00 00 00"                      '//=> | mov eax,fs:[00000030]
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H7, "05 00 10 00 00"                            '//+7 | add eax,1000
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &HC, "8B 40 20"                                  '//+7 | mov eax,[eax+20]
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &HF, "83 C0 68"                                  '//+A | add eax,68
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H12, "8B 5C 24 04"                              '//+D | mov ebx,[esp+04]
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H16, "89 18"                                    '//+11 | mov [eax],ebx
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H18, "89 44 24 04"                              '//+13 | mov [esp+04],eax
        FuckWriteAOBByString hProcess, ByVal BaseAddress + &H1C, "C2 10 00"                                 '//+17 | ret 0010
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + &H1F, ByVal StrPtr(lpImagePath), sLenth, False  '//+1A | inc esp ¼g¤JImagePathName
       
        CloseHandleAPI hProcess
        CallWindowProc BaseAddress, BaseAddress + &H1F, 0, 0, 0
End If

        
        BeUsed = True
        ByPassHs = BaseAddress

    End If
End Function
Public Function FuckWriteAOBByString(ByVal FuckHandle As Long, ByVal lpAddress As Long, ByVal lpAobString As String) As Long
    Dim WriteI As Integer
    Dim WriteStr() As String
    Dim WriteBuff() As Byte
    lpAobString = Trim(lpAobString)
    WriteStr() = Split(lpAobString, " ")
    ReDim WriteBuff(UBound(WriteStr))
    For WriteI = 0 To UBound(WriteStr)
        WriteBuff(WriteI) = (Val("&H" + WriteStr(WriteI)))
    Next WriteI
    FuckWriteAOBByString = WriteProcessMemoryAPI(FuckHandle, ByVal lpAddress, ByVal VarPtr(WriteBuff(0)), UBound(WriteBuff) + 1, 0&)
End Function
Public Function InjectDll(ByVal pszLibFile As String) As Long
    If Not Inited Then Exit Function
    Dim hProcess As Long, hThread As Long
    Dim pszLibFileRemote As Long, exitCode As Long
    On Error GoTo errhandle
    hProcess = OpenProcessAPI(PROCESS_QUERY_INFORMATION Or PROCESS_CREATE_THREAD Or PROCESS_VM_OPERATION Or PROCESS_VM_WRITE, 0, Pid)
    If hProcess = 0 Then GoTo errhandle
    Dim cch   As Long, cb As Long
    cch = 1 + LenB(StrConv(pszLibFile, vbFromUnicode))
    cb = cch
    pszLibFileRemote = VirtualAllocEx(hProcess, ByVal 0&, cb, MEM_COMMIT, PAGE_READWRITE)
    If pszLibFileRemote = 0 Then GoTo errhandle
    If (WriteProcessMemoryAPI(hProcess, ByVal pszLibFileRemote, ByVal pszLibFile, cb, ByVal 0&) = 0) Then GoTo errhandle
    Dim pfnThreadRtn As Long
    pfnThreadRtn = GetProcAddress(GetModuleHandle("Kernel32"), "LoadLibraryA")
    If pfnThreadRtn = 0 Then GoTo errhandle
    hThread = CreateRemoteThread(hProcess, ByVal 0&, 0&, ByVal pfnThreadRtn, ByVal pszLibFileRemote, 0, 0&)
    If (hThread = 0) Then GoTo errhandle
    WaitForSingleObject hThread, 10000
    GetExitCodeThread hThread, exitCode
    InjectDll = exitCode
    Exit Function
errhandle:
    If pszLibFileRemote <> 0 Then
        VirtualFreeEx hProcess, ByVal pszLibFileRemote, 0, MEM_RELEASE
        InjectDll = 0
        Exit Function
    End If
    If hThread <> 0 Then
        CloseHandleAPI hThread
        InjectDll = 0
        Exit Function
    End If
    If hProcess <> 0 Then
        CloseHandleAPI hProcess
        InjectDll = 0
        Exit Function
    End If
End Function

Public Function WriteMemory(ByVal lpAddress As Long, ByVal lpBuffer As Long, ByVal lpSize As Long) As Long
WriteMemory = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal lpBuffer, ByVal lpSize, False)
End Function

Public Function WriteByte(ByVal lpAddress As Long, ByVal lpValue As Byte) As Long
    WriteByte = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteInteger(ByVal lpAddress As Long, ByVal lpValue As Integer) As Long
    WriteInteger = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteLong(ByVal lpAddress As Long, ByVal lpValue As Long) As Long
    WriteLong = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteCurrency(ByVal lpAddress As Long, ByVal lpValue As Currency) As Long
    WriteCurrency = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteSingle(ByVal lpAddress As Long, ByVal lpValue As Single) As Long
    WriteSingle = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteDouble(ByVal lpAddress As Long, ByVal lpValue As Double) As Long
    WriteDouble = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteString(ByVal lpAddress As Long, ByVal lpValue As String) As Long
    WriteString = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(lpValue), ByVal LenB(lpValue), False)
End Function

Public Function WriteAOBByString(ByVal lpAddress As Long, ByVal lpAobString As String) As Long
    Dim WriteI As Integer
    Dim WriteStr() As String
    Dim WriteBuff() As Byte
    lpAobString = Trim(lpAobString)
    WriteStr() = Split(lpAobString, " ")
    ReDim WriteBuff(UBound(WriteStr))
    For WriteI = 0 To UBound(WriteStr)
        WriteBuff(WriteI) = (Val("&H" + WriteStr(WriteI)))
    Next WriteI
    WriteAOBByString = WriteProcessMemoryAPI(handle, ByVal lpAddress, ByVal VarPtr(WriteBuff(0)), UBound(WriteBuff) + 1, 0&)
End Function

Public Function WritePointer(ByVal lpAddress As Long, ByVal lpOffset As Long, ByVal lpValue As Long) As Long
    WritePointer = WriteLong(ReadLong(lpAddress) + lpOffset, lpValue)
End Function

Public Function WriteMultiPointerByString(ByVal lpPointerList As String, ByVal lpBuffer As Long, ByVal lpSize As Long) As Long
'HACK.WriteMultiPointerByString "5A3B08=>460:A8:0", VarPtr(Gold), 4
    Dim PointerList() As String
    Dim tmp1, tmp2, i As Long
    tmp1 = Val("&H" & Split(lpPointerList, "=>")(0))
    PointerList = Split(Split(lpPointerList, "=>")(1), ":")
    For i = 0 To UBound(PointerList)
        tmp2 = ReadLong(tmp1) + Val("&H" & PointerList(i))
        tmp1 = tmp2
    Next
    tmp2 = WriteMemory(tmp1, lpBuffer, lpSize)
    WriteMultiPointerByString = tmp2
End Function

Public Function ReadMemory(ByVal lpAddress As Long, ByVal lpBuffer As Long, ByVal lpSize As Long) As Long
ReadMemory = ReadProcessMemoryAPI(handle, ByVal lpAddress, ByVal lpBuffer, ByVal lpSize, False)
End Function

Public Function ReadByte(ByVal lpAddress As Long) As Byte
    Dim Value As Byte
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 1, False
    ReadByte = Value
End Function

Public Function ReadInteger(ByVal lpAddress As Long) As Integer
    Dim Value As Integer
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 2, False
    ReadInteger = Value
End Function

Public Function ReadLong(ByVal lpAddress As Long) As Long
    Dim Value As Long
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 4, False
    ReadLong = Value
End Function

Public Function ReadCurrency(ByVal lpAddress As Long) As Currency
    Dim Value As Currency
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 8, False
    ReadCurrency = Value
End Function

Public Function ReadSingle(ByVal lpAddress As Long) As Single
    Dim Value As Single
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 4, False
    ReadSingle = Value
End Function

Public Function ReadDouble(ByVal lpAddress As Long) As Double
    Dim Value As Double
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 8, False
    ReadDouble = Value
End Function

Public Function ReadString(ByVal lpAddress As Long, ByVal lpSize As Long) As String
    Dim Value As String
    Value = Space(lpSize)
    ReadProcessMemoryAPI handle, ByVal lpAddress, ByVal StrPtr(Value), ByVal lpSize, False
    ReadString = Value
    Value = ""
End Function

Public Function ReadPointer(ByVal lpAddress As Long, ByVal lpOffset As Long) As Long
    ReadPointer = ReadLong(ReadLong(lpAddress) + lpOffset)
End Function

Public Function ReadMultiPointerByString(ByVal lpPointerList As String, ByVal lpBuffer As Long, ByVal lpSize As Long) As Long
'HACK.ReadMultiPointerByString "5A3B08=>460:A8:C", VarPtr(Gold), 4
    Dim PointerList() As String
    Dim tmp1, tmp2, i As Long
    tmp1 = Val("&H" & Split(lpPointerList, "=>")(0))
    PointerList = Split(Split(lpPointerList, "=>")(1), ":")
    For i = 0 To UBound(PointerList)
        tmp2 = ReadLong(tmp1) + Val("&H" & PointerList(i))
        tmp1 = tmp2
    Next
    tmp2 = ReadMemory(tmp1, lpBuffer, lpSize)
    ReadMultiPointerByString = tmp2
End Function

Public Function Alloc(ByVal lpSize As Long, Optional ByVal lpAddress As Long = 0) As Long
    If Not Inited Then Exit Function
    If PreAllocAddress = 0 Then
        PreAllocAddress = VirtualAllocEx(handle, ByVal lpAddress, ByVal PreAllocSize, MEM_COMMIT, PAGE_READWRITE)
    End If
    If lpSize > (PreAllocSize / 4) Then
        Alloc = VirtualAllocEx(handle, ByVal lpAddress, ByVal lpSize, MEM_COMMIT, PAGE_READWRITE)
        Exit Function
    End If
    If lpSize + AllocedSize > PreAllocSize Then
        PreAllocAddress = VirtualAllocEx(handle, ByVal lpAddress, ByVal PreAllocSize, MEM_COMMIT, PAGE_READWRITE)
        AllocedSize = 0
    End If
    Alloc = PreAllocAddress + AllocedSize
    AllocedSize = AllocedSize + lpSize
End Function

Public Function MakeJmp(ByVal lpAddress As Long, ByVal lpJmpAddress As Long, Optional ByVal lpNops As Long = 0) As Long
    MakeJmp = CBool(WriteByte(lpAddress, &HE9)) And CBool(WriteLong(lpAddress + 1, lpJmpAddress - lpAddress - 5))
    If lpNops = 0 Then Exit Function
    MakeJmp = MakeJmp And CBool(MakeNops(lpAddress + 5, lpNops))
End Function

Public Function MakeCall(ByVal lpAddress As Long, ByVal lpCallAddress As Long, Optional ByVal lpNops As Long = 0) As Long
    MakeCall = CBool(WriteByte(lpAddress, &HE8)) And CBool(WriteLong(lpAddress + 1, lpCallAddress - lpAddress - 5))
    If lpNops = 0 Then Exit Function
    MakeCall = MakeCall And CBool(MakeNops(lpAddress + 5, lpNops))
End Function

Public Function MakeNops(ByVal lpAddress As Long, ByVal lpSize As Long) As Long
    If lpSize = 0 Then
        MakeNops = True
        Exit Function
    End If
    Dim NOP() As Byte
    ReDim NOP(lpSize)
    Dim i As Long
    For i = 0 To lpSize - 1
        NOP(i) = &H90
    Next
    MakeNops = WriteMemory(lpAddress, VarPtr(NOP(0)), lpSize)
End Function

Public Function GetFunctionAddress(ByVal lpModule As String, ByVal lpFunction As String) As Long
    GetFunctionAddress = GetProcAddress(LoadLibrary(lpModule), lpFunction)
End Function
Public Function SEND_MESSAGE(nMessage As String, Optional ByVal KeyDown As Integer = 1)
  Dim data() As Byte, i As Long
  i = 0

  data = StrConv(nMessage, vbFromUnicode)

  While i <= UBound(data)
      If data(i) < 128 Then
          PostMessage hWnd, &H102, data(i), 0&
          i = i + 1
      Else
          PostMessage hWnd, &H102, data(i), 0&
          PostMessage hWnd, &H102, data(i + 1), 0&
          i = i + 2
      End If
  Wend
If KeyDown = 1 Then PostMessage hWnd, &H100, vbKeyReturn, 0&
End Function
Public Function SEND_KEYDOWN(ByVal nKeys As Long)
   Call PostMessage(hWnd, &H100, nKeys, 0&)
   Call PostMessage(hWnd, &H101, nKeys, 0&)
End Function
Public Function tounsign(ByVal x As Long) As Double
tounsign = CDbl("&H" & Left(Hex(x), 4)) * 65536 + CDbl("&H" & Right(Hex(x), 4))
End Function


