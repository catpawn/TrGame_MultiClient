Attribute VB_Name = "Module1"
Private Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "Psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function RtlAdjustPrivilege Lib "ntdll" (ByVal Privilege As Long, ByVal bEnablePrivilege As Long, ByVal bCurrentThread As Long, ByRef OldState As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByVal Dst As Long, ByVal Src As Long, ByVal uLen As Long)

Private Declare Function NtQueryInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByVal ProcessInformation As Long, ByVal ProcessInformationLength As Long, ReturnLength As Long) As Long
Public Declare Function ZwQueryInformationThread Lib "Ntdll.dll " (ByVal hThread As Long, ByVal ThreadInformationClass As Long, ByRef ThreadInformation As Long, ByVal ThreadInformationLength As Long, ReturnLength As Long) As Long
Public Declare Function ZwQueryInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As PROCESSINFOCLASS, ByVal ProcessInformation As Long, ByVal ProcessInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function GetMappedFileName Lib "Psapi.dll" Alias "GetMappedFileNameA" (ByVal hProcess As Long, ByVal lpv As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
 Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Type LUID
   lowpart As Long
   highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges As LUID_AND_ATTRIBUTES
End Type


Type CLIENT_ID
        UniqueProcess As Long
        UniqueThread As Long
End Type

Type THREAD_BASIC_INFORMATION
    ExitStatus As Long
    TebBaseAddress As Long
    ClientId As CLIENT_ID
    AffinityMask As Long
    PRIORITY As Long
    BasePriority As Long
End Type



Public Enum PROCESSINFOCLASS
ProcessBasicInformation
ProcessQuotaLimits
ProcessIoCounters
ProcessVmCounters
ProcessTimes
ProcessBasePriority
ProcessRaisePriority
ProcessDebugPort
ProcessExceptionPort
ProcessAccessToken
ProcessLdtInformation
ProcessLdtSize
ProcessDefaultHardErrorMode
ProcessIoPortHandlers '// Note: this is kernel mode only
ProcessPooledUsageAndLimits
ProcessWorkingSetWatch
ProcessUserModeIOPL
ProcessEnableAlignmentFaultFixup
ProcessPriorityClass
ProcessWx86Information
ProcessHandleCount
ProcessAffinityMask
ProcessPriorityBoost
ProcessDeviceMap
ProcessSessionInformation
ProcessForegroundInformation
ProcessWow64Information
ProcessImageFileName
ProcessLUIDDeviceMapsEnabled
ProcessBreakOnTermination
ProcessDebugObjectHandle
ProcessDebugFlags
ProcessHandleTracing
ProcessIoPriority
ProcessExecuteFlags
ProcessResourceManagement
ProcessCookie
ProcessImageInformation
MaxProcessInfoClass
End Enum




Public Type PROCESSENTRY32
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
Public Type THREADENTRY32
  dwSize As Long
  cntUsage As Long
  th32ThreadID As Long
  th32OwnerProcessID As Long
  tpBasePri As Long
  tpDeltaPri As Long
  dwFlags As Long
End Type

Public Const THREAD_SUSPEND_RESUME As Long = &H2

Public ProcessId() As PROCESSENTRY32
Public Thread() As THREADENTRY32

Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_INHERIT = &H80000000
Public Const TH32CS_SNAPall = (TH32CS_SNAPheaplist Or TH32CS_SNAPPROCESS Or TH32CS_SNAPthread Or TH32CS_SNAPmodule)
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ           As Long = 16&
Private Const PROCESS_VM_WRITE          As Long = (&H20)
Private Const PROCESS_VM_OPERATION      As Long = (&H8)
Private Const MAX_PATH                  As Long = 260

Public Type PROCESS_BASIC_INFORMATION
    ExitStatus                      As Long
    PebBaseAddress                  As Long
    AffinityMask                    As Long
    BasePriority                    As Long
    UniqueProcessId                 As Long
    InheritedFromUniqueProcessId    As Long
End Type

Private Type LIST_ENTRY
    Blink                           As Long
    Flink                           As Long
End Type

Private Type UNICODE_STRING
    Length                          As Long
    MaximumLength                   As Long
    Buffer                          As Long
End Type

Private Type LDR_MODULE 'LDR_DATA_TABLE_ENTRY
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
    BaseAddress                     As Long
    EntryPoint                      As Long
    SizeOfImage                     As Long
    FullDllName                     As UNICODE_STRING
    BaseDllName                     As UNICODE_STRING
    Flags                           As Long
    LoadCount                       As Integer
    TlsIndex                        As Integer
    HashTableEntry                  As LIST_ENTRY
    TimeDateStamp                   As Long
End Type

Private Type PEB_LDR_DATA
    Length                          As Long
    Initialized                     As Long
    SsHandle                        As Long
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
End Type
Public Function GetPidPID(ProcessName As String) As Long
    Dim objWMIService, objProcess, colProcess
    Dim strComputer
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process")

    For Each objProcess In colProcess

        If objProcess.Name = ProcessName Then
           GetPidPID = objProcess.ProcessId

            Exit For

        End If

    Next

    Set objWMIService = Nothing
    Set colProcess = Nothing

End Function



Public Function Thread32_Enum(ByRef Thread() As THREADENTRY32, lProcessID As Long) As Long
ReDim Thread(0)
Dim THREADENTRY32 As THREADENTRY32
Dim hSnapshot As Long
Dim lThread As Long
hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPthread, lProcessID)
THREADENTRY32.dwSize = Len(THREADENTRY32)
If Thread32First(hSnapshot, THREADENTRY32) = False Then
  Thread32_Enum = -1
  Exit Function
Else
  ReDim Thread(lThread)
  Thread(lThread) = THREADENTRY32
End If

Do
  If Thread32Next(hSnapshot, THREADENTRY32) = False Then
  Exit Do
  Else
  lThread = lThread + 1
  ReDim Preserve Thread(lThread)
  Thread(lThread) = THREADENTRY32
  End If
Loop
Thread32_Enum = lThread
End Function

Function Thread_Suspend(T_ID As Long) As Long
Dim hThread As Long
hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
Thread_Suspend = SuspendThread(hThread)
End Function

Function Thread_Resume(T_ID As Long) As Long
Dim hThread As Long

Dim lSuspendCount As Long
hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
Thread_Resume = ResumeThread(hThread)
End Function

Public Sub TList()

End Sub

Private Function GetPEB(ByVal lPid As Long) As Long
    Dim tPBI    As PROCESS_BASIC_INFORMATION
    Dim lRet    As Long
    Dim lProc   As Long
    lProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lPid)
    If lProc Then
        If NtQueryInformationProcess(lProc, 0, VarPtr(tPBI), Len(tPBI), lRet) = 0 Then
            GetPEB = tPBI.PebBaseAddress
        End If
        CloseHandle lProc
    End If
End Function

Function GetImageNameByThread(ProcessHandle As Long, ByVal PHandle As Long) As String
 Dim TBI As THREAD_BASIC_INFORMATION
    Dim STATUS As Long
     Dim hThread As Long
   Dim hProcess As Long
   Dim StartAddr As Long
   Dim ImageName As String * 250
   Dim trm As String
   ZwQueryInformationThread PHandle, 9, StartAddr, Len(StartAddr), 0
      GetMappedFileName ProcessHandle, ByVal StartAddr, ImageName, Len(ImageName)
      trm = Replace(ImageName, Chr(32), "")
      
      If Len(Trim(trm)) > 0 Then

       GetImageNameByThread = trm
      End If
   
End Function


Public Function EnablePrivilege(seName As String) As Boolean
    Dim p_lngRtn As Long
    Dim p_lngToken As Long
    Dim p_lngBufferLen As Long
    Dim p_typLUID As LUID
    Dim p_typTokenPriv As TOKEN_PRIVILEGES
    Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
    p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
    If p_lngRtn = 0 Then
        Exit Function ' Failed
    ElseIf Err.LastDllError <> 0 Then
        Exit Function ' Failed
    End If
    p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)  'Used to look up privileges LUID.
    If p_lngRtn = 0 Then
        Exit Function ' Failed
    End If
    ' Set it up to adjust the program's security privilege.
    p_typTokenPriv.PrivilegeCount = 1
    p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
    p_typTokenPriv.Privileges.pLuid = p_typLUID
    EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function

 Function 只获取数字(nStr1 As String) As String
  Dim i As Long, nStr As String, mStr As String, Str1 As String
  nStr = nStr1
  For i = 1 To Len(nStr)
     Str1 = Mid(nStr, i, 1)
     If IsNumeric(Str1) Then mStr = mStr & Str1
  Next
 只获取数字 = mStr
End Function
