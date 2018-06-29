Attribute VB_Name = "Module2"
Public Declare Function InternetOpen Lib "Wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "Wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetReadFile Lib "Wininet.dll" (ByVal hFile As Long, sBuffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "Wininet.dll" (ByVal hInet As Long) As Integer

Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_FLAG_RELOAD = &H80000000
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

Public Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Public Sub Delay(DelayTime As Single)
Dim ST As Single
  ST = Timer
  Do Until Timer - ST > DelayTime
    DoEvents
  Loop
End Sub
Function OpenURL(ByVal sUrl As String) As String
    On Error Resume Next
    Dim hOpen As Long
    Dim hOpenUrl As Long
    Dim bDoLoop As Boolean
    Dim bRet As Boolean
    Dim sReadBuffer As String
    Dim lNumberOfBytesRead As Long
    Dim sBuffer As String
    hOpen = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    bDoLoop = True
    While bDoLoop
        sReadBuffer = Space$(2048)
        bRet = InternetReadFile(hOpenUrl, ByVal sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = Trim(sBuffer)
End Function
Public Function GetHTML(HTML, ByVal Pattern As String) As String
    Dim strData As String
    Dim reg As Object
    Dim matchs As Object, match As Object
    strData = HTML
    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = False
    reg.Pattern = Pattern
    Set matchs = reg.Execute(strData)
    For Each match In matchs
        GetHTML = match.SubMatches(0)
        Next
End Function
Function UTF8ToUrl(str) As String
On Error Resume Next
Dim GetBytes() As Byte
Dim retStr As String
retStr = ""
With CreateObject("ADODB.Stream")
.Mode = 3
.Type = 2
.Open
.Charset = "UTF-8"
.WriteText (str)
.Position = 0
.Type = 1
GetBytes = .Read(-1)
.Close
End With
    For i = 3 To UBound(GetBytes)
        retStr = retStr & "%" & Hex(GetBytes(i))
    Next
    UTF8ToUrl = retStr
End Function
