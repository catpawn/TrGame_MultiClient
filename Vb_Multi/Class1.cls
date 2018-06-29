VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, lpStructure As Any, ByVal dwSize As Long, ByVal dwReserved As Long) As Long

Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As Currency          'Idle
    dwSpare(0 To 75) As Long
End Type

Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type

Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As Currency
    liKeSystemTime As Currency
    liExpTimeZoneBias  As Currency
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private lidOldIdle As Currency
Private liOldSystem As Currency
Private sbSysBasicInfo As SYSTEM_BASIC_INFORMATION

Public Function GetCPUUsage() As Long
    Dim spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim stSysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim curIdle As Currency
    Dim curSystem As Currency
    GetCPUUsage = -1
    If NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, stSysTimeInfo, Len(stSysTimeInfo), 0&) <> NO_ERROR Then Exit Function
    If NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, spSysPerforfInfo, Len(spSysPerforfInfo), 0&) <> NO_ERROR Then Exit Function

    curIdle = spSysPerforfInfo.liIdleTime - lidOldIdle
    curSystem = stSysTimeInfo.liKeSystemTime - liOldSystem
   
    If curSystem <> 0 Then curIdle = (curIdle / curSystem) * 100
    GetCPUUsage = CLng(100 - curIdle / sbSysBasicInfo.bKeNumberProcessors)
    lidOldIdle = spSysPerforfInfo.liIdleTime
    liOldSystem = stSysTimeInfo.liKeSystemTime
End Function

Private Sub Class_Initialize()
    Dim stSysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION
    If NtQuerySystemInformation(SYSTEM_BASICINFORMATION, sbSysBasicInfo, Len(sbSysBasicInfo), 0&) <> NO_ERROR Then Exit Sub
    If NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, stSysTimeInfo, Len(stSysTimeInfo), 0&) <> NO_ERROR Then Exit Sub
    If NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, spSysPerforfInfo, LenB(spSysPerforfInfo), ByVal 0&) <> NO_ERROR Then Exit Sub
    lidOldIdle = spSysPerforfInfo.liIdleTime
    liOldSystem = stSysTimeInfo.liKeSystemTime
End Sub
