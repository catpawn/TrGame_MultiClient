VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "In BlackTr We Trust"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7560
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   6480
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打開遊戲"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      OLEDropMode     =   1  '手動
      TabIndex        =   6
      Top             =   2040
      Width           =   3735
   End
   Begin VB.ListBox List3 
      Height          =   600
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5520
      Top             =   2400
   End
   Begin VB.ListBox List2 
      Height          =   780
      Left            =   3960
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   2400
   End
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "雙開破解"
      Height          =   300
      Left            =   6360
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "遊戲窗口"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "進程名稱"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "遊戲路徑:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hack As New cls_hack
Dim FirstGamePid As Long
Dim FirstRun As Boolean
Private Sub Check1_Click()
Timer1.Enabled = Check1.Value
End Sub



Private Function EnumAllProcess()
List1.Clear
Dim i As Integer
Dim theloop As Long, snap As Long
Dim Proc As PROCESSENTRY32
snap = CreateToolhelp32Snapshot(TH32CS_SNAPall, 0)
Proc.dwSize = Len(Proc)
theloop = Process32First(snap, Proc)
  While theloop <> 0
  ReDim Preserve ProcessId(i)
  ProcessId(i).th32ProcessID = Proc.th32ProcessID
  ProcessId(i).cntThreads = Proc.cntThreads
  ProcessId(i).th32ParentProcessID = Proc.th32ParentProcessID
  List1.AddItem "Pid=" & Proc.th32ProcessID & "," & "exeName=" & Proc.szexeFile
  i = i + 1
  theloop = Process32Next(snap, Proc)
  Wend
CloseHandle snap
End Function
Private Function GetFirstPid()
If FirstRun = True Then
If List2.ListCount <> 0 Then
FirstGamePid = Replace(GetHTML(List2.List(0), "Pid=(.+,)"), ",", "")
hack.OpenProcess FirstGamePid
hack.InjectDll App.Path & "/lpk.dll"
hack.CloseHandle
FirstRun = False
Timer3.Enabled = True
Timer2.Enabled = False
End If
End If
End Function
Private Function GetNotFirstGamePid()
List3.Clear
For i = 0 To List2.ListCount - 1
If InStr(List2.List(i), FirstGamePid) = 0 Then
List3.AddItem Replace(GetHTML(List2.List(i), "Pid=(.+,)"), ",", "")
End If
Next i
End Function
Private Function EnableMulti()
For i = 0 To List3.ListCount - 1
If List3.ListCount <> 0 Then
hack.OpenProcess List3.List(i)
Delay 1.5
hack.WriteAOBByString &H4B33F7, "90 90 90 90 90 90"
hack.WriteAOBByString &HBC0E19, "74"
hack.WriteAOBByString &HB216EC, "0F 85"
hack.InjectDll App.Path & "/lpk.dll"
hack.CloseHandle
End If
Next i
End Function

Private Sub Command2_Click()
Shell "cmd.exe /c start " & Chr(34) & Chr(34) & " " & Chr(34) & Text1.Text & Chr(34) & " -p:2 -zone:lvs.talesrunner.com.hk -zoneID:1 -zoneName:HK_Test -- hongkong"
End Sub

Private Sub Form_Load()
hack.ByPassHs
FirstRun = True
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If InStr(Data.Files(1), ".lnk") Then
MsgBox "這是一個捷徑檔案！", vbExclamation, "提示"
Else
If InStr(Data.Files(1), "trgame.exe") Then
Text1 = Data.Files(1)
Text1.Enabled = False
Else
MsgBox "這不是一個有效的Talesrunner執行檔案！", vbInformation, "提示"
End If
End If
End Sub

Private Sub Timer1_Timer()
EnumAllProcess
List2.Clear
ListView1.ListItems.Clear
For i = 0 To List1.ListCount - 1
 If InStr(List1.List(i), "trgame.exe") Then
 List2.AddItem List1.List(i)
 ListView1.ListItems.Add , , Replace(GetHTML(List1.List(i), "Pid=(.+,)"), ",", "")
ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "trgame.exe"
 End If
 Next i
End Sub

Private Sub Timer2_Timer()
GetFirstPid
End Sub

Private Sub Timer3_Timer()
GetNotFirstGamePid
EnableMulti
End Sub

Private Sub Timer4_Timer()
If InStr(Text1.Text, "trgame.exe") Then
Command2.Enabled = True
Else
Command2.Enabled = False
End If
End Sub
