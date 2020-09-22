VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpeedXP"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTips 
      Caption         =   "Tips"
      Height          =   315
      Left            =   7793
      TabIndex        =   26
      Top             =   3525
      Width           =   825
   End
   Begin VB.Frame Frame3 
      Caption         =   "Help"
      Height          =   1785
      Left            =   165
      TabIndex        =   19
      Top             =   3300
      Width           =   7275
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Move Your Cursor To Any Item To Get Information About It"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   135
         TabIndex        =   20
         Top             =   285
         Width           =   6960
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   7605
      TabIndex        =   3
      Top             =   4695
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Network - Internet"
      Height          =   3030
      Left            =   4560
      TabIndex        =   2
      Top             =   195
      Width           =   4155
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   24
         Text            =   "8760"
         Top             =   645
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":57E2
         Left            =   2505
         List            =   "frmMain.frx":57F2
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2580
         Width           =   1410
      End
      Begin VB.CheckBox Check15 
         Caption         =   "IE Max Connections Per Server"
         Height          =   255
         Left            =   165
         TabIndex        =   18
         Top             =   2220
         Width           =   2925
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Network Adapter Onboard Processor"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   1905
         Width           =   2925
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Optimize Max Duplicate ACKs"
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   1590
         Width           =   2925
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Enable Selective Acknowledgement"
         Height          =   255
         Left            =   165
         TabIndex        =   15
         Top             =   1275
         Width           =   2925
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Enable MTU Auto Discovery"
         Height          =   255
         Left            =   165
         TabIndex        =   14
         Top             =   960
         Width           =   2925
      End
      Begin VB.CheckBox Check10 
         Caption         =   "TCP Window Size (RWIN)"
         Height          =   255
         Left            =   165
         TabIndex        =   13
         Top             =   645
         Width           =   2925
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Large TCP Window (RWIN) Support"
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   330
         Width           =   2925
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1110
         Picture         =   "frmMain.frx":5825
         Top             =   2595
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Recommended"
         Height          =   315
         Left            =   1380
         TabIndex        =   25
         Top             =   2625
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System (File System/Cache)"
      Height          =   3030
      Left            =   165
      TabIndex        =   1
      Top             =   195
      Width           =   4155
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3195
         MaxLength       =   4
         TabIndex        =   23
         Top             =   1905
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3195
         MaxLength       =   6
         TabIndex        =   21
         Top             =   330
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Clear PageFile at Shutdown"
         Height          =   255
         Left            =   165
         TabIndex        =   11
         Top             =   2535
         Width           =   2925
      End
      Begin VB.CheckBox Check7 
         Caption         =   "CMOS/RealTimeClock Priority Boost"
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   2220
         Width           =   2925
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Explicitly Specify L2 Cache"
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   1905
         Width           =   2925
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Enable UDMA66 on Intel Chipsets"
         Height          =   255
         Left            =   165
         TabIndex        =   8
         Top             =   1590
         Width           =   2925
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Disable NTFS Last Access Update"
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   1275
         Width           =   2925
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Disable 8.3 Filename Creation"
         Height          =   255
         Left            =   165
         TabIndex        =   6
         Top             =   960
         Width           =   2925
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Disable NT Executive Paging"
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   645
         Width           =   2925
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Optimize I/O Page Lock Limit"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   330
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   360
      Left            =   7605
      TabIndex        =   0
      Top             =   4215
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*  Author:     Asim Aziz                               *
'*  Date:       Apr 2, 2002                             *
'*  Mail:       chirisoft.flashmail.com                 *
'*  URL:        http://chirisoft.cjb.net                *
'*                                                      *
'*  This program uses windows hidden registry settings  *
'*  to optimize the general system and network          *
'*  performance according to the hardware you own.      *
'*  Please Compile the file with same name as the       *
'*  manifest (or rename the manifest file) to see visual*
'*  themes.                                             *
'********************************************************


Option Explicit
Dim Changed As Boolean

Private Sub Form_Initialize()
  InitCommonControls 'initialize common controls for XP themes to work
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Dim hKey As Long
  Dim lValue As Long
  
  'registry values (if exist) are read from registry.
  
  'open registry key for querying a value.
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management", 0, KEY_QUERY_VALUE, hKey
  
  lValue = 0
  RegQueryValueEx hKey, "IOPageLockLimit", 0&, REG_DWORD, lValue, 4
  If lValue <> 0 Then
    Check1.Value = Checked
    Text1 = lValue
  End If
  
  lValue = 0
  RegQueryValueEx hKey, "DisablePagingExecutive", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check2.Value = Checked
  
  lValue = 0
  RegQueryValueEx hKey, "ClearPageFileAtShutdown", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check8.Value = Checked
  
  lValue = 0
  RegQueryValueEx hKey, "SecondLevelDataCache", 0&, REG_DWORD, lValue, 4
  If lValue <> 0 Then
    Check6.Value = Checked
    Text2 = lValue
  End If
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", 0, KEY_QUERY_VALUE, hKey
  
  lValue = 0
  RegQueryValueEx hKey, "NtfsDisable8dot3NameCreation", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check3.Value = Checked
  
  lValue = 0
  RegQueryValueEx hKey, "NtfsDisableLastAccessUpdate", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check4.Value = Checked
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\PriorityControl", 0, KEY_QUERY_VALUE, hKey
  
  lValue = 0
  RegQueryValueEx hKey, "IRQ8Priority", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check7.Value = Checked
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\Class\{4D36E96A-E325-11CE-BFC1-08002BE10318}\0000", 0, KEY_QUERY_VALUE, hKey
  
  lValue = 0
  RegQueryValueEx hKey, "EnableUDMA66", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check5.Value = Checked
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_QUERY_VALUE, hKey
  
  lValue = 0
  RegQueryValueEx hKey, "Tcp1323Opts", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check9.Value = Checked
  
  lValue = 0
  RegQueryValueEx hKey, "GlobalMaxTcpWindowSize", 0&, REG_DWORD, lValue, 4
  If lValue <> 0 Then
    Check10.Value = Checked
    Text3 = lValue
  End If
  
  lValue = 0
  RegQueryValueEx hKey, "EnablePMTUDiscovery", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check11.Value = Checked
  
  lValue = 0
  RegQueryValueEx hKey, "SackOpts", 0&, REG_DWORD, lValue, 4
  If lValue = 1 Then Check12.Value = Checked
  
  lValue = 0
  RegQueryValueEx hKey, "TcpMaxDupAcks", 0&, REG_DWORD, lValue, 4
  If lValue = 2 Then Check13.Value = Checked
  
  lValue = 0
  If RegQueryValueEx(hKey, "DisableTaskOffload", 0&, REG_DWORD, lValue, 4) = 0 Then
    If lValue = 0 Then Check14.Value = Checked
  End If
  
  RegCloseKey hKey
  
  
  
  
  
  RegOpenKeyEx HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", 0, KEY_QUERY_VALUE, hKey
  
  lValue = 0
  RegQueryValueEx hKey, "MaxConnectionsPerServer", 0&, REG_DWORD, lValue, 4
  If lValue <> 0 Then Check15.Value = Checked
  
  RegCloseKey hKey
End Sub

Private Sub cmdApply_Click()
  On Error Resume Next
  Dim hKey As Long
  
  'Registry values are written to the registry.
  
  Changed = True
  
  'Open a registry key for setting a value.
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management", 0, KEY_SET_VALUE, hKey
  
  If Check1.Value = Checked Then
    RegSetValueEx hKey, "IOPageLockLimit", 0, REG_DWORD, CLng(Text1), 4
  Else
    RegSetValueEx hKey, "IOPageLockLimit", 0, REG_DWORD, 0, 4
  End If
  
  If Check2.Value = Checked Then
    RegSetValueEx hKey, "DisablePagingExecutive", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "DisablePagingExecutive", 0, REG_DWORD, 0, 4
  End If
  
  If Check8.Value = Checked Then
    RegSetValueEx hKey, "ClearPageFileAtShutdown", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "ClearPageFileAtShutdown", 0, REG_DWORD, 0, 4
  End If
  
  If Check6.Value = Checked Then
    RegSetValueEx hKey, "SecondLevelDataCache", 0, REG_DWORD, CLng(Text2), 4
  End If
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", 0, KEY_SET_VALUE, hKey
  
  If Check3.Value = Checked Then
    RegSetValueEx hKey, "NtfsDisable8dot3NameCreation", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "NtfsDisable8dot3NameCreation", 0, REG_DWORD, 0, 4
  End If
  
  If Check4.Value = Checked Then
    RegSetValueEx hKey, "NtfsDisableLastAccessUpdate", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "NtfsDisableLastAccessUpdate", 0, REG_DWORD, 0, 4
  End If
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\PriorityControl", 0, KEY_SET_VALUE, hKey
  
  If Check7.Value = Checked Then
    RegSetValueEx hKey, "IRQ8Priority", 0, REG_DWORD, 1, 4
  Else
    RegDeleteValue hKey, "IRQ8Priority"
  End If
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\Class\{4D36E96A-E325-11CE-BFC1-08002BE10318}\0000", 0, KEY_SET_VALUE, hKey
  
  If Check5.Value = Checked Then
    RegSetValueEx hKey, "EnableUDMA66", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "EnableUDMA66", 0, REG_DWORD, 0, 4
  End If
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_SET_VALUE, hKey
  
  If Check9.Value = Checked Then
    RegSetValueEx hKey, "Tcp1323Opts", 0, REG_DWORD, 1, 4
  Else
    RegDeleteValue hKey, "Tcp1323Opts"
  End If
  
  If Check10.Value = Checked Then
    RegSetValueEx hKey, "GlobalMaxTcpWindowSize", 0, REG_DWORD, CLng(Text3), 4
    RegSetValueEx hKey, "TcpWindowSize", 0, REG_DWORD, CLng(Text3), 4
  Else
    RegDeleteValue hKey, "GlobalMaxTcpWindowSize"
    RegDeleteValue hKey, "TcpWindowSize"
  End If
  
  If Check11.Value = Checked Then
    RegSetValueEx hKey, "EnablePMTUDiscovery", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "EnablePMTUDiscovery", 0, REG_DWORD, 0, 4
  End If
  
  If Check12.Value = Checked Then
    RegSetValueEx hKey, "SackOpts", 0, REG_DWORD, 1, 4
  Else
    RegSetValueEx hKey, "SackOpts", 0, REG_DWORD, 0, 4
  End If
  
  If Check13.Value = Checked Then
    RegSetValueEx hKey, "TcpMaxDupAcks", 0, REG_DWORD, 2, 4
  Else
    RegDeleteValue hKey, "TcpMaxDupAcks"
  End If
  
  If Check14.Value = Checked Then
    RegSetValueEx hKey, "DisableTaskOffload", 0, REG_DWORD, 0, 4
  Else
    RegDeleteValue hKey, "DisableTaskOffload"
  End If
  
  RegCloseKey hKey
  
  
  
  
  RegOpenKeyEx HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", 0, KEY_SET_VALUE, hKey
  
  If Check15.Value = Checked Then
    RegSetValueEx hKey, "MaxConnectionsPerServer", 0, REG_DWORD, 20, 4
    RegSetValueEx hKey, "MaxConnectionsPer1_0Server", 0, REG_DWORD, 20, 4
  Else
    RegDeleteValue hKey, "MaxConnectionsPerServer"
    RegDeleteValue hKey, "MaxConnectionsPer1_0Server"
  End If
  
  RegCloseKey hKey
End Sub






'Show the help when user moves the mouse over a control
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Move Your Cursor To Any Item To Get Information About It"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Move Your Cursor To Any Item To Get Information About It"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Move Your Cursor To Any Item To Get Information About It"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Move Your Cursor To Any Item To Get Information About It"
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Diskcache plays a very important role in WinXP. However, the default I/O pagefile setting of XP is conservative, which limits the performance. Some better values for different RAM are given below." & vbCrLf & vbCrLf & "64M: 4096   (4M)" & vbCrLf & "128M: 16384    (16M)" & vbCrLf & "256M: 65536     (64M)" & vbCrLf & "512M or more: 262144     (256M)"
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "On systems with large amount of RAM this tweak can be enabled to force the core Windows system to be kept in memory and not paged to disk." & vbCrLf & vbCrLf & "To make sure that you will benifit from this tweak run all the applications that you normally do simultaneously." & vbCrLf & "Check ""Task Manager"" for free physical memory." & vbCrLf & "If you still have 64M+ free, this tweak is for you."
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "By default, NTFS also generates the style of file name that consists of eight characters, followed by a period and a three-character extension for compatibility with MS-DOS and Microsoft® Windows® 3.x clients. If you are not supporting these types of clients, you can safely turn off this setting to improve NTFS performance."
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "By default NTFS updates the date and time stamp of the last access on files and directories whenever a file or directory is accessed. For a large NTFS volume, this update process can slow performance."
End Sub

Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "If you have a computer with an Intel chipset that supports UDMA66 (810 and later), It is disabled by default in Windows XP. This tweak allows you to enable or disable it" & vbCrLf & vbCrLf & "Note: Before you enable UDMA66 mode make sure that the device supports UDMA66 mode and use an 80-pin IDE cable with the proper pin cut. "
End Sub

Private Sub Check6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Windows Doesn't do a very good job at detecting processor L2 Cache. It is incorrectly detected most of the times. If you know exactly how much L2 cache your processor has, specefy it here in KiloBytes"
End Sub

Private Sub Check7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "In windows XP you can prioritise any of the 15 IRQs to impove the responsiveness of the device on that IRQ, but CMOS/RealTimeClock (at IRQ8) gives overall system performance gain."
End Sub

Private Sub Check8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "While it doesn't give any performance gains but You may enable it for security reasons" & vbCrLf & vbCrLf & "Caution -- PC Shutdown will slow down as Pagefile has to be cleared"
End Sub

Private Sub Check9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Enables Large TCPWindow support as described in RFC 1323. Without this parameter, the TCPWindow is limited to 64K"
End Sub

Private Sub Check10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "For best results, the TCPWindow should be a multiple of MSS (Maximum Segment Size). MSS is generally MTU - 40, where MTU (Maximum Transmission Unit) is the largest packet size that can be transmitted. MTU is usually 1500 (1492 for PPPoE connections)." & vbCrLf & "You may have to experiment with different values to find that perfect setting for your specific setup."
End Sub

Private Sub Check11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "When enabled, TCP attempts to discover MTU  automatically over the path to a remote host. Setting this parameter to False causes MTU to default to 576 which reduces overall performance over high speed connections. "
End Sub

Private Sub Check12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Controls whether or not Selective ACK (SACK - RFC 2018) support is enabled. With SACK enabled (default), a packet or series of packets can be dropped, and the receiver informs the sender which data has been received, and where there may be ""holes"" in the data.So only the packets not received are sent instead of whole window. This is specially important if you are using large TCP window sizes."
End Sub

Private Sub Check13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "Optimize the number of duplicate ACKs that must be received for the same sequence number of sent data before ""fast retransmit"" is triggered to resend the segment that has been dropped in transit"
End Sub

Private Sub Check14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "If your network adapter has an onboard processor, designed to offload network processing from the system CPU, it is disabled by default. This setting allows you to enable it and increase the processing speed of your system. "
End Sub

Private Sub Check15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label1 = "According to the HTTP specs, only limited number of simultaneous connections are allowed, while loading pages. This setting optimizes the number of connections. " & vbCrLf & "Note: Keep in mind that although this setting works fine in most cases, it exceeds the HTTP specs and therefore might cause problems with some websites. If you experience problems, just disable it. While this setting might improve web page loading considerably, it tends to strain webservers more and has no effect on throughput.  (It is a Per User setting)"
End Sub





'Restrict the textboxes to accept only numeric values
Private Sub Text1_KeyPress(KeyAscii As Integer)
  If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub





Private Sub Check1_Click()
  If Check1.Value = Checked Then Text1.Visible = True Else Text1.Visible = False
End Sub

Private Sub Check6_Click()
  If Check6.Value = Checked Then Text2.Visible = True Else Text2.Visible = False
End Sub

Private Sub Check9_Click()
  Combo1.ListIndex = 0
End Sub

Private Sub Check10_Click()
  Combo1.ListIndex = 0
  If Check10.Value = Checked Then Text3.Visible = True Else Text3.Visible = False
End Sub

Private Sub Check11_Click()
  Combo1.ListIndex = 0
End Sub

Private Sub Check12_Click()
  Combo1.ListIndex = 0
End Sub

Private Sub Check13_Click()
  Combo1.ListIndex = 0
End Sub

Private Sub Check14_Click()
  Combo1.ListIndex = 0
End Sub

Private Sub Check15_Click()
  Combo1.ListIndex = 0
End Sub



'Load recommended values for network settings
Private Sub Combo1_Click()
If Combo1.ListIndex = 1 Then
  Check9.Value = Checked
  Check10.Value = Checked
  Text3 = 256960
  Check11.Value = Checked
  Check12.Value = Checked
  Check13.Value = Checked
  Check14.Value = Checked
  Check15.Value = Checked
  
ElseIf Combo1.ListIndex = 2 Then
  Check9.Value = Checked
  Check10.Value = Checked
  Text3 = 255552
  Check11.Value = Checked
  Check12.Value = Checked
  Check13.Value = Checked
  Check14.Value = Checked
  Check15.Value = Checked
  
ElseIf Combo1.ListIndex = 3 Then
  Check9.Value = Unchecked
  Check10.Value = Checked
  Text3 = 5840
  Check11.Value = Checked
  Check12.Value = Checked
  Check13.Value = Checked
  Check14.Value = Unchecked
  Check15.Value = Checked
  
End If
End Sub

Private Sub cmdTips_Click()
frmTips.Show vbModal, frmMain
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Changed = True Then
  SHRestartSystemMB Me.hWnd, vbNullString, 2 Or 4 'show restart dialogue
End If
End Sub

