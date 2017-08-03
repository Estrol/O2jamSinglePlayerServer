VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "版权HOHO所有"
   ClientHeight    =   5565
   ClientLeft      =   765
   ClientTop       =   630
   ClientWidth     =   7680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "关于"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   7335
      Begin MSWinsockLib.Winsock wskhttp 
         Index           =   0
         Left            =   960
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   15000
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   315
         Left            =   5280
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "my name is hoho, my qq is 2460739."
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2100
         TabIndex        =   7
         Top             =   300
         Width           =   3195
      End
   End
   Begin MSWinsockLib.Winsock wsk 
      Index           =   0
      Left            =   7140
      Top             =   4740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15010
   End
   Begin VB.Frame Frame2 
      Caption         =   "服务控制"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "傻瓜型启动劲乐团个人服务器版"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "退出"
         Height          =   375
         Left            =   1740
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "复位服务器"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "网络交换信息"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   2895
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "Form1.frx":57E2
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim spass(54) As Byte
Dim fun(30) As Integer
Dim p9(8) As Byte, p10(9) As Byte, p8(7) As Byte, p4(3) As Byte, p12(11) As Byte, loadserver(54) As Byte, p16(15) As Byte, p6(5) As Byte, p5(4) As Byte
Dim room(3907) As Byte, data1(1049) As Byte, userinfo(309) As Byte, data2(79) As Byte, data3(4425) As Byte, createroom(12) As Byte, jz() As Byte
Dim finish(75) As Byte, msg(72) As Byte
Dim version As String, httpmsg1 As String, httpmsg2 As String, httpheader As String

Private Function enters()
Dim i As Integer
Dim str As String
str = InputBox("输入服务器IP地址或者计算机名称", "服务器连接", "127.8.8.8")
If Len(str) = 0 Then Exit Function
i = Shell("otwo 1 o2jam-patchbj02.9you.com o2jam/patch  " & str & ":15000 1 1 1 1 1 1 1 1 " & str & " 15010 " & str & " 15010 " & str & " 15010 " & str & " 15010 " & str & " 15010", vbNormalFocus)
End Function
Private Sub Command1_Click()
If Dir("OTwo.exe", vbArchive + vbNormal) = "" Then
            MsgBox "把服务器程序和劲乐团放在一个目录下傻瓜型启动,谢谢!", 4096 + vbCritical, "HOHO Help System"
            Exit Sub
End If
Dim i As Integer
i = Shell("OTwo.exe 1 o2jam-patchbj02.9you.com o2jam/patch 127.8.8.8:15000 1 1 1 1 1 1 1 127.8.8.8 15010 127.8.8.8 15010 127.8.8.8 15010 127.8.8.8 15010 127.8.8.8 15010 127.8.8.8 15010 127.8.8.8 15010", vbNormalFocus)
End Sub

Private Sub Command2_Click()
Call resetwsk
Text1.Text = "Copyright HOHO"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Dim i As Integer
For i = 1 To 5
            If wsk(i).State = sckConnected Then
                        wsk(i).SendData msg()
                        addsock i, "接收帮助信息"
            End If
Next
End Sub

Private Sub Form_Load()
On Error GoTo xXx
version = "测试版 V0.20"
Me.Caption = "O2Jam Personal Server " & version
Text1.Text = Text1.Text & vbCrLf & version
'If Date > DateAdd("s", 1, "2005-5-3") Then
'            MsgBox "测试期限到了！不过你可以调整系统时间到2005年5月3号或更早继续使用。", 4096 + vbApplicationModal + vbExclamation, version
'            End
'End If
Dim i As Integer
Dim ii As Integer

httpmsg1 = "//Lightboard" & vbCrLf & _
"欢迎您来到HOHO劲乐单机版世界！" & vbCrLf & vbCrLf & _
"//Stateroom" & vbCrLf & _
"抵制不良游戏,拒绝盗版游戏,注意自我保护,谨防上当受骗" & vbCrLf & _
"适度游戏益脑,沉迷游戏伤身,合理安排时间,享受健康生活" & vbCrLf & _
"对于单机版任何疑问或者问题,请联系QQ2460739" & vbCrLf & vbCrLf

httpmsg2 = "//StateWaiting" & vbCrLf & _
"欢迎使用HOHO劲乐团单机版服务器程序!" & vbCrLf & _
"F2                 : 自动更换颜色" & vbCrLf & _
"F3                 : 快速开始.单机模式没有房主 " & vbCrLf & _
"F5                 : 游戏途中切换音符类型" & vbCrLf & _
"F6                 : 游戏途中调整音符帮助" & vbCrLf & _
"F7                 : 切换音符显示模式2D / 3D" & vbCrLf & _
"F8                 : 切换鼠标模式window/image" & vbCrLf & _
"F9                 : 均衡器转换模式 开 / 关" & vbCrLf & _
"PrintScreen        : 截图"

httpheader = "HTTP/1.1 200 OK" & vbCrLf & _
"Server: O2Jam PS/0.10" & vbCrLf & _
"Date: Sat, 19 June 2004 00:00:00 GMT" & vbCrLf & _
"Content-Type: text/plain" & vbCrLf & _
"Accept-Ranges: bytes" & vbCrLf & _
"Last-Modified: Sat, 19 June 2004 00:00:00 GMT" & vbCrLf & _
"Content-Length: "

msg(0) = &H49
msg(1) = &H0
msg(2) = &H7C
msg(3) = &HE6
msg(4) = &H84
msg(5) = &H9F
msg(6) = &HE8
msg(7) = &HB0
msg(8) = &HA2
msg(9) = &HE4
msg(10) = &HBD
msg(11) = &HA0
msg(12) = &HE5
msg(13) = &HAF
msg(14) = &HB9
msg(15) = Asc("H")
msg(16) = Asc("O")
msg(17) = Asc("H")
msg(18) = Asc("O")
msg(19) = &HE8
msg(20) = &HBD
msg(21) = &HAF
msg(22) = &HE4
msg(23) = &HBB
msg(24) = &HB6
msg(25) = &HE7
msg(26) = &H9A
msg(27) = &H84
msg(28) = &HE6
msg(29) = &H94
msg(30) = &HAF
msg(31) = &HE6
msg(32) = &H8C
msg(33) = &H81
msg(34) = &HEF
msg(35) = &HBC
msg(36) = &H8C
msg(37) = &HE5
msg(38) = &HBD
msg(39) = &H93
msg(40) = &HE5
msg(41) = &H89
msg(42) = &H8D
msg(43) = &HE6
msg(44) = &H9C
msg(45) = &H8D
msg(46) = &HE5
msg(47) = &H8A
msg(48) = &HA1
msg(49) = &HE5
msg(50) = &H99
msg(51) = &HA8
msg(52) = &HE7
msg(53) = &H89
msg(54) = &H88
msg(55) = &HE6
msg(56) = &H9C
msg(57) = &HAC
msg(58) = Asc("V")
msg(59) = Asc("0")
msg(60) = Asc(".")
msg(61) = Asc("2")
msg(62) = Asc("0")
msg(63) = &HE6
msg(64) = &HB5
msg(65) = &H8B
msg(66) = &HE8
msg(67) = &HAF
msg(68) = &H95
msg(69) = &HE7
msg(70) = &H89
msg(71) = &H88
msg(72) = &H0

finish(0) = &H4C
finish(2) = &HB2
finish(3) = &HF
finish(4) = &H8
finish(9) = &H1
finish(29) = &HD0
finish(31) = &H3
finish(35) = &HB4
finish(36) = &HB
finish(41) = &H1
finish(46) = &H2
finish(51) = &H3
finish(56) = &H4
finish(61) = &H5
finish(66) = &H6
finish(71) = &H7

spass(0) = &H37
spass(1) = &H0
spass(2) = &HF2
spass(3) = &H3

p5(0) = 5
p6(0) = 6
p16(0) = 16
p4(0) = 4
p8(0) = 8
p10(0) = 10
p9(0) = 9
p12(0) = 12

createroom(0) = &HD
createroom(2) = &HD6
createroom(3) = &H7
createroom(8) = &H3

data3(0) = &H4A
data3(1) = &H11
data3(2) = &HD3
data3(3) = &H7
data3(4) = &HC8
data3(12) = &H2
data3(13) = &H5B
data3(14) = &H77
data3(15) = &H7A
data3(16) = &H2E
data3(17) = &H5D
data3(18) = &H6B
data3(19) = &H73
data3(20) = &H5F
data3(21) = &HC5
data3(22) = &HC0
data3(23) = &HB5
data3(24) = &HAF
data3(25) = &HD7
data3(26) = &HE0
data3(27) = &HB7
data3(28) = &HBF
data3(29) = &HBC
data3(30) = &HE4
data3(32) = &H1
data3(33) = &HDA
data3(34) = &HC0
data3(35) = &H2
data3(36) = &H1
data3(37) = &H4
data3(38) = &H8
data3(39) = &H2
data3(48) = &H1
data3(58) = &HFF
data3(70) = &H2
data3(80) = &HFF
data3(92) = &H3
data3(102) = &HFF
data3(114) = &H4
data3(124) = &HFF
data3(136) = &H5
data3(146) = &HFF
data3(158) = &H6
data3(168) = &HFF
data3(180) = &H7
data3(190) = &HFF
data3(202) = &H8
data3(212) = &HFF
data3(224) = &H9
data3(234) = &HFF
data3(246) = &HA
data3(256) = &HFF
data3(268) = &HB
data3(278) = &HFF
data3(290) = &HC
data3(300) = &HFF
data3(312) = &HD
data3(322) = &HFF
data3(334) = &HE
data3(344) = &HFF
data3(356) = &HF
data3(366) = &HFF
data3(378) = &H10
data3(388) = &HFF
data3(400) = &H11
data3(410) = &HFF
data3(422) = &H12
data3(432) = &HFF
data3(444) = &H13
data3(454) = &HFF
data3(466) = &H14
data3(476) = &HFF
data3(488) = &H15
data3(498) = &HFF
data3(510) = &H16
data3(520) = &HFF
data3(532) = &H17
data3(542) = &HFF
data3(554) = &H18
data3(564) = &HFF
data3(576) = &H19
data3(586) = &HFF
data3(598) = &H1A
data3(608) = &HFF
data3(620) = &H1B
data3(630) = &HFF
data3(642) = &H1C
data3(652) = &HFF
data3(664) = &H1D
data3(674) = &HFF
data3(686) = &H1E
data3(696) = &HFF
data3(708) = &H1F
data3(718) = &HFF
data3(730) = &H20
data3(740) = &HFF
data3(752) = &H21
data3(762) = &HFF
data3(774) = &H22
data3(784) = &HFF
data3(796) = &H23
data3(806) = &HFF
data3(818) = &H24
data3(828) = &HFF
data3(840) = &H25
data3(850) = &HFF
data3(862) = &H26
data3(872) = &HFF
data3(884) = &H27
data3(894) = &HFF
data3(906) = &H28
data3(916) = &HFF
data3(928) = &H29
data3(938) = &HFF
data3(950) = &H2A
data3(960) = &HFF
data3(972) = &H2B
data3(982) = &HFF
data3(994) = &H2C
data3(1004) = &HFF
data3(1016) = &H2D
data3(1026) = &HFF
data3(1038) = &H2E
data3(1048) = &HFF
data3(1060) = &H2F
data3(1070) = &HFF
data3(1082) = &H30
data3(1092) = &HFF
data3(1104) = &H31
data3(1114) = &HFF
data3(1126) = &H32
data3(1136) = &HFF
data3(1148) = &H33
data3(1158) = &HFF
data3(1170) = &H34
data3(1180) = &HFF
data3(1192) = &H35
data3(1202) = &HFF
data3(1214) = &H36
data3(1224) = &HFF
data3(1236) = &H37
data3(1246) = &HFF
data3(1258) = &H38
data3(1268) = &HFF
data3(1280) = &H39
data3(1290) = &HFF
data3(1302) = &H3A
data3(1312) = &HFF
data3(1324) = &H3B
data3(1334) = &HFF
data3(1346) = &H3C
data3(1356) = &HFF
data3(1368) = &H3D
data3(1378) = &HFF
data3(1390) = &H3E
data3(1400) = &HFF
data3(1412) = &H3F
data3(1422) = &HFF
data3(1434) = &H40
data3(1444) = &HFF
data3(1456) = &H41
data3(1466) = &HFF
data3(1478) = &H42
data3(1488) = &HFF
data3(1500) = &H43
data3(1510) = &HFF
data3(1522) = &H44
data3(1532) = &HFF
data3(1544) = &H45
data3(1554) = &HFF
data3(1566) = &H46
data3(1576) = &HFF
data3(1588) = &H47
data3(1598) = &HFF
data3(1610) = &H48
data3(1620) = &HFF
data3(1632) = &H49
data3(1642) = &HFF
data3(1654) = &H4A
data3(1664) = &HFF
data3(1676) = &H4B
data3(1686) = &HFF
data3(1698) = &H4C
data3(1708) = &HFF
data3(1720) = &H4D
data3(1730) = &HFF
data3(1742) = &H4E
data3(1752) = &HFF
data3(1764) = &H4F
data3(1774) = &HFF
data3(1786) = &H50
data3(1796) = &HFF
data3(1808) = &H51
data3(1818) = &HFF
data3(1830) = &H52
data3(1840) = &HFF
data3(1852) = &H53
data3(1862) = &HFF
data3(1874) = &H54
data3(1884) = &HFF
data3(1896) = &H55
data3(1906) = &HFF
data3(1918) = &H56
data3(1928) = &HFF
data3(1940) = &H57
data3(1950) = &HFF
data3(1962) = &H58
data3(1972) = &HFF
data3(1984) = &H59
data3(1994) = &HFF
data3(2006) = &H5A
data3(2016) = &HFF
data3(2028) = &H5B
data3(2038) = &HFF
data3(2050) = &H5C
data3(2060) = &HFF
data3(2072) = &H5D
data3(2082) = &HFF
data3(2094) = &H5E
data3(2104) = &HFF
data3(2116) = &H5F
data3(2126) = &HFF
data3(2138) = &H60
data3(2148) = &HFF
data3(2160) = &H61
data3(2170) = &HFF
data3(2182) = &H62
data3(2192) = &HFF
data3(2204) = &H63
data3(2214) = &HFF
data3(2226) = &H64
data3(2236) = &HFF
data3(2248) = &H65
data3(2258) = &HFF
data3(2270) = &H66
data3(2280) = &HFF
data3(2292) = &H67
data3(2302) = &HFF
data3(2314) = &H68
data3(2324) = &HFF
data3(2336) = &H69
data3(2346) = &HFF
data3(2358) = &H6A
data3(2368) = &HFF
data3(2380) = &H6B
data3(2390) = &HFF
data3(2402) = &H6C
data3(2412) = &HFF
data3(2424) = &H6D
data3(2434) = &HFF
data3(2446) = &H6E
data3(2456) = &HFF
data3(2468) = &H6F
data3(2478) = &HFF
data3(2490) = &H70
data3(2500) = &HFF
data3(2512) = &H71
data3(2522) = &HFF
data3(2534) = &H72
data3(2544) = &HFF
data3(2556) = &H73
data3(2566) = &HFF
data3(2578) = &H74
data3(2588) = &HFF
data3(2600) = &H75
data3(2610) = &HFF
data3(2622) = &H76
data3(2632) = &HFF
data3(2644) = &H77
data3(2654) = &HFF
data3(2666) = &H78
data3(2676) = &HFF
data3(2688) = &H79
data3(2698) = &HFF
data3(2710) = &H7A
data3(2720) = &HFF
data3(2732) = &H7B
data3(2742) = &HFF
data3(2754) = &H7C
data3(2764) = &HFF
data3(2776) = &H7D
data3(2786) = &HFF
data3(2798) = &H7E
data3(2808) = &HFF
data3(2820) = &H7F
data3(2830) = &HFF
data3(2842) = &H80
data3(2852) = &HFF
data3(2864) = &H81
data3(2874) = &HFF
data3(2886) = &H82
data3(2896) = &HFF
data3(2908) = &H83
data3(2918) = &HFF
data3(2930) = &H84
data3(2940) = &HFF
data3(2952) = &H85
data3(2962) = &HFF
data3(2974) = &H86
data3(2984) = &HFF
data3(2996) = &H87
data3(3006) = &HFF
data3(3018) = &H88
data3(3028) = &HFF
data3(3040) = &H89
data3(3050) = &HFF
data3(3062) = &H8A
data3(3072) = &HFF
data3(3084) = &H8B
data3(3094) = &HFF
data3(3106) = &H8C
data3(3116) = &HFF
data3(3128) = &H8D
data3(3138) = &HFF
data3(3150) = &H8E
data3(3160) = &HFF
data3(3172) = &H8F
data3(3182) = &HFF
data3(3194) = &H90
data3(3204) = &HFF
data3(3216) = &H91
data3(3226) = &HFF
data3(3238) = &H92
data3(3248) = &HFF
data3(3260) = &H93
data3(3270) = &HFF
data3(3282) = &H94
data3(3292) = &HFF
data3(3304) = &H95
data3(3314) = &HFF
data3(3326) = &H96
data3(3336) = &HFF
data3(3348) = &H97
data3(3358) = &HFF
data3(3370) = &H98
data3(3380) = &HFF
data3(3392) = &H99
data3(3402) = &HFF
data3(3414) = &H9A
data3(3424) = &HFF
data3(3436) = &H9B
data3(3446) = &HFF
data3(3458) = &H9C
data3(3468) = &HFF
data3(3480) = &H9D
data3(3490) = &HFF
data3(3502) = &H9E
data3(3512) = &HFF
data3(3524) = &H9F
data3(3534) = &HFF
data3(3546) = &HA0
data3(3556) = &HFF
data3(3568) = &HA1
data3(3578) = &HFF
data3(3590) = &HA2
data3(3600) = &HFF
data3(3612) = &HA3
data3(3622) = &HFF
data3(3634) = &HA4
data3(3644) = &HFF
data3(3656) = &HA5
data3(3666) = &HFF
data3(3678) = &HA6
data3(3688) = &HFF
data3(3700) = &HA7
data3(3710) = &HFF
data3(3722) = &HA8
data3(3732) = &HFF
data3(3744) = &HA9
data3(3754) = &HFF
data3(3766) = &HAA
data3(3776) = &HFF
data3(3788) = &HAB
data3(3798) = &HFF
data3(3810) = &HAC
data3(3820) = &HFF
data3(3832) = &HAD
data3(3842) = &HFF
data3(3854) = &HAE
data3(3864) = &HFF
data3(3876) = &HAF
data3(3886) = &HFF
data3(3898) = &HB0
data3(3908) = &HFF
data3(3920) = &HB1
data3(3930) = &HFF
data3(3942) = &HB2
data3(3952) = &HFF
data3(3964) = &HB3
data3(3974) = &HFF
data3(3986) = &HB4
data3(3996) = &HFF
data3(4008) = &HB5
data3(4018) = &HFF
data3(4030) = &HB6
data3(4040) = &HFF
data3(4052) = &HB7
data3(4062) = &HFF
data3(4074) = &HB8
data3(4084) = &HFF
data3(4096) = &HB9
data3(4106) = &HFF
data3(4118) = &HBA
data3(4128) = &HFF
data3(4140) = &HBB
data3(4150) = &HFF
data3(4162) = &HBC
data3(4172) = &HFF
data3(4184) = &HBD
data3(4194) = &HFF
data3(4206) = &HBE
data3(4216) = &HFF
data3(4228) = &HBF
data3(4238) = &HFF
data3(4250) = &HC0
data3(4260) = &HFF
data3(4272) = &HC1
data3(4282) = &HFF
data3(4294) = &HC2
data3(4304) = &HFF
data3(4316) = &HC3
data3(4326) = &HFF
data3(4338) = &HC4
data3(4348) = &HFF
data3(4360) = &HC5
data3(4370) = &HFF
data3(4382) = &HC6
data3(4392) = &HFF
data3(4404) = &HC7
data3(4414) = &HFF


data2(0) = &H50
data2(2) = &HDB
data2(3) = &H7
data2(4) = &H3
data2(8) = &H79
data2(9) = &H75
data2(10) = &H62
data2(11) = &H69
data2(12) = &H61
data2(13) = &H6F
data2(14) = &H32
data2(15) = &H31
data2(16) = &H39
data2(18) = &H5B
data2(19) = &H57
data2(20) = &H7A
data2(21) = &H2E
data2(22) = &H5D
data2(23) = &H59
data2(24) = &H42
data2(25) = &H5F
data2(26) = &HC2
data2(27) = &HE4
data2(28) = &HB2
data2(29) = &HDD
data2(31) = &H7
data2(35) = &H77
data2(36) = &H69
data2(37) = &H6E
data2(38) = &H72
data2(39) = &H61
data2(40) = &H69
data2(41) = &H6E
data2(42) = &H38
data2(43) = &H33
data2(45) = &H5B
data2(46) = &H77
data2(47) = &H7A
data2(48) = &H2E
data2(49) = &H5D
data2(50) = &H6B
data2(51) = &H73
data2(52) = &H5F
data2(53) = &HC5
data2(54) = &HC0
data2(56) = &HD
data2(60) = &H71
data2(61) = &H7A
data2(62) = &H6A
data2(63) = &H68
data2(64) = &H6F
data2(65) = &H68
data2(66) = &H6F
data2(68) = &H71
data2(69) = &H7A
data2(70) = &H6A
data2(71) = &H68
data2(72) = &H6F
data2(73) = &H68
data2(74) = &H6F
data2(76) = &HE


userinfo(0) = &H36
userinfo(1) = &H1
userinfo(2) = &HD1
userinfo(3) = &H7
userinfo(8) = &H71
userinfo(9) = &H7A
userinfo(10) = &H6A
userinfo(11) = &H68
userinfo(12) = &H6F
userinfo(13) = &H68
userinfo(14) = &H6F
userinfo(16) = &H1
userinfo(17) = &HC
userinfo(18) = &H72
userinfo(29) = &HE
userinfo(33) = &H7E
userinfo(37) = &H27
userinfo(41) = &H57
userinfo(45) = &H40
userinfo(46) = &H7C
userinfo(54) = &H3D
userinfo(55) = &H1
userinfo(58) = &H11
userinfo(59) = &H1
userinfo(62) = &H3E
userinfo(63) = &H1
userinfo(66) = &H36
userinfo(70) = &H33
userinfo(74) = &H5B
userinfo(78) = &HC1
userinfo(82) = &H3C
userinfo(94) = &HFF
userinfo(98) = &H23
userinfo(118) = &H99
userinfo(122) = &H9B
userinfo(154) = &H93
userinfo(158) = &H95
userinfo(162) = &H9D
userinfo(166) = &H97
userinfo(245) = &H5
userinfo(258) = &H6
userinfo(262) = &H93
userinfo(266) = &H14
userinfo(270) = &H95
userinfo(274) = &H14
userinfo(278) = &H97
userinfo(282) = &H14
userinfo(286) = &H99
userinfo(290) = &HA
userinfo(294) = &H9B
userinfo(298) = &HA
userinfo(302) = &H9D
userinfo(306) = &H14

data1(0) = &H1A
data1(1) = &H4
data1(2) = &HBF
data1(3) = &HF
data1(4) = &H57
data1(6) = &H64
data1(8) = &H79
data1(9) = &H1
data1(10) = &H34
data1(11) = &H2
data1(12) = &H5B
data1(13) = &H2
data1(18) = &H66
data1(20) = &H3E
data1(21) = &H1
data1(22) = &H41
data1(23) = &H2
data1(24) = &H19
data1(25) = &H3
data1(30) = &H67
data1(32) = &HA4
data1(34) = &H2E
data1(35) = &H1
data1(36) = &HF8
data1(37) = &H1
data1(42) = &H68
data1(44) = &H45
data1(45) = &H1
data1(46) = &HA2
data1(47) = &H1
data1(48) = &H27
data1(49) = &H2
data1(54) = &H69
data1(56) = &HD6
data1(58) = &H87
data1(59) = &H1
data1(60) = &HE8
data1(61) = &H1
data1(66) = &H6A
data1(68) = &HA8
data1(70) = &H64
data1(71) = &H1
data1(72) = &HCF
data1(73) = &H1
data1(78) = &H6B
data1(80) = &H48
data1(81) = &H1
data1(82) = &HB7
data1(83) = &H1
data1(84) = &H7D
data1(85) = &H3
data1(90) = &H6C
data1(92) = &HD3
data1(94) = &H25
data1(95) = &H1
data1(96) = &H92
data1(97) = &H1
data1(102) = &H6D
data1(104) = &H22
data1(105) = &H1
data1(106) = &H5C
data1(107) = &H1
data1(108) = &H54
data1(109) = &H2
data1(114) = &H6E
data1(116) = &H59
data1(117) = &H1
data1(118) = &H4A
data1(119) = &H2
data1(120) = &HF1
data1(121) = &H3
data1(126) = &H6F
data1(128) = &H26
data1(129) = &H1
data1(130) = &H2B
data1(131) = &H2
data1(132) = &H24
data1(133) = &H4
data1(138) = &H70
data1(140) = &H1C
data1(141) = &H1
data1(142) = &H3F
data1(143) = &H2
data1(144) = &H14
data1(145) = &H3
data1(150) = &H71
data1(152) = &H3
data1(153) = &H1
data1(154) = &H31
data1(155) = &H2
data1(156) = &H20
data1(157) = &H3
data1(162) = &H72
data1(164) = &HE2
data1(166) = &H34
data1(167) = &H2
data1(168) = &HF1
data1(169) = &H2
data1(174) = &H77
data1(176) = &H87
data1(177) = &H4
data1(178) = &H87
data1(179) = &H4
data1(180) = &H87
data1(181) = &H4
data1(186) = &H79
data1(188) = &H38
data1(189) = &H1
data1(190) = &HF0
data1(191) = &H2
data1(192) = &HE0
data1(193) = &H4
data1(198) = &H7C
data1(200) = &H98
data1(202) = &H4E
data1(203) = &H1
data1(204) = &HBF
data1(205) = &H1
data1(210) = &H7E
data1(212) = &H40
data1(213) = &H1
data1(214) = &HCD
data1(215) = &H1
data1(216) = &H7D
data1(217) = &H2
data1(222) = &H80
data1(224) = &H45
data1(225) = &H2
data1(226) = &HBB
data1(227) = &H3
data1(228) = &HF9
data1(229) = &H4
data1(234) = &H82
data1(236) = &H9
data1(237) = &H1
data1(238) = &HAA
data1(239) = &H1
data1(240) = &H7
data1(241) = &H2
data1(246) = &H83
data1(248) = &H7E
data1(250) = &H27
data1(251) = &H1
data1(252) = &H86
data1(253) = &H1
data1(258) = &H84
data1(260) = &H2F
data1(261) = &H1
data1(262) = &HD3
data1(263) = &H2
data1(264) = &HC7
data1(265) = &H5
data1(270) = &H85
data1(272) = &HCE
data1(273) = &H1
data1(274) = &H7C
data1(275) = &H2
data1(276) = &HF3
data1(277) = &H2
data1(282) = &H86
data1(284) = &H70
data1(285) = &H1
data1(286) = &HEA
data1(287) = &H1
data1(288) = &H9D
data1(289) = &H2
data1(294) = &H88
data1(296) = &H48
data1(297) = &H1
data1(298) = &H1A
data1(299) = &H2
data1(300) = &HB5
data1(301) = &H2
data1(306) = &H8A
data1(308) = &HDE
data1(310) = &H5A
data1(311) = &H2
data1(312) = &H1
data1(313) = &H3
data1(318) = &H8B
data1(320) = &H6A
data1(321) = &H1
data1(322) = &HF4
data1(323) = &H1
data1(324) = &H12
data1(325) = &H3
data1(330) = &H8D
data1(332) = &H37
data1(333) = &H1
data1(334) = &H6
data1(335) = &H2
data1(336) = &H83
data1(337) = &H2
data1(342) = &H91
data1(344) = &H89
data1(346) = &HDE
data1(348) = &H44
data1(349) = &H1
data1(354) = &H92
data1(356) = &H4
data1(357) = &H1
data1(358) = &H8F
data1(359) = &H1
data1(360) = &HD8
data1(361) = &H1
data1(366) = &H93
data1(368) = &H6C
data1(370) = &HCE
data1(372) = &H3C
data1(373) = &H1
data1(378) = &H94
data1(380) = &HFB
data1(382) = &HE5
data1(383) = &H1
data1(384) = &H80
data1(385) = &H3
data1(390) = &H95
data1(392) = &H3D
data1(393) = &H1
data1(394) = &H12
data1(395) = &H2
data1(396) = &HD4
data1(397) = &H4
data1(402) = &H96
data1(404) = &H1
data1(405) = &H1
data1(406) = &H53
data1(407) = &H2
data1(408) = &HEA
data1(409) = &H2
data1(414) = &H97
data1(416) = &H2F
data1(417) = &H1
data1(418) = &HBB
data1(419) = &H1
data1(420) = &H60
data1(421) = &H2
data1(426) = &H98
data1(428) = &HF9
data1(430) = &H96
data1(431) = &H1
data1(432) = &HF0
data1(433) = &H1
data1(438) = &H9F
data1(440) = &HF5
data1(442) = &H4F
data1(443) = &H1
data1(444) = &HB6
data1(445) = &H1
data1(450) = &HA0
data1(452) = &H21
data1(453) = &H1
data1(454) = &HE6
data1(455) = &H1
data1(456) = &H73
data1(457) = &H2
data1(462) = &HA1
data1(464) = &HDC
data1(466) = &HFE
data1(467) = &H1
data1(468) = &H8D
data1(469) = &H2
data1(474) = &HA2
data1(476) = &H1F
data1(477) = &H1
data1(478) = &HAD
data1(479) = &H1
data1(480) = &H3D
data1(481) = &H2
data1(486) = &HA3
data1(488) = &H11
data1(489) = &H1
data1(490) = &HF8
data1(491) = &H1
data1(492) = &HF
data1(493) = &H3
data1(498) = &HA4
data1(500) = &H3B
data1(501) = &H1
data1(502) = &HEF
data1(503) = &H1
data1(504) = &H54
data1(505) = &H2
data1(510) = &HA5
data1(512) = &HCF
data1(514) = &HFE
data1(515) = &H1
data1(516) = &HAD
data1(517) = &H2
data1(522) = &HA6
data1(524) = &HF0
data1(526) = &H2A
data1(527) = &H1
data1(528) = &HE0
data1(529) = &H1
data1(534) = &HA7
data1(536) = &H58
data1(538) = &HE
data1(539) = &H1
data1(540) = &H5A
data1(541) = &H1
data1(546) = &HA8
data1(548) = &H7
data1(549) = &H1
data1(550) = &HB4
data1(551) = &H1
data1(552) = &H6F
data1(553) = &H2
data1(558) = &HAA
data1(560) = &H97
data1(561) = &H1
data1(562) = &HAE
data1(563) = &H2
data1(564) = &HC3
data1(565) = &H4
data1(570) = &HAC
data1(572) = &H63
data1(573) = &H1
data1(574) = &H90
data1(575) = &H2
data1(576) = &HD
data1(577) = &H3
data1(582) = &HAD
data1(584) = &H38
data1(585) = &H1
data1(587) = &H2
data1(588) = &H11
data1(589) = &H3
data1(594) = &HAE
data1(596) = &HD1
data1(598) = &HD4
data1(599) = &H1
data1(600) = &H9F
data1(601) = &H2
data1(606) = &HAF
data1(608) = &HAB
data1(610) = &HE7
data1(611) = &H1
data1(612) = &H8E
data1(613) = &H2
data1(618) = &HB0
data1(620) = &HEB
data1(622) = &HF7
data1(623) = &H1
data1(624) = &HCE
data1(625) = &H2
data1(630) = &HB2
data1(632) = &H59
data1(633) = &H1
data1(634) = &H19
data1(635) = &H2
data1(636) = &H78
data1(637) = &H2
data1(642) = &HBA
data1(644) = &HE8
data1(645) = &H1
data1(646) = &HFC
data1(647) = &H2
data1(648) = &H51
data1(649) = &H4
data1(654) = &HBC
data1(656) = &H45
data1(657) = &H1
data1(658) = &H80
data1(659) = &H2
data1(660) = &H93
data1(661) = &H3
data1(666) = &HBD
data1(668) = &HE2
data1(670) = &H8B
data1(671) = &H1
data1(672) = &H30
data1(673) = &H2
data1(678) = &HBF
data1(680) = &H67
data1(681) = &H1
data1(682) = &H3E
data1(683) = &H2
data1(684) = &HFB
data1(685) = &H3
data1(690) = &HC1
data1(692) = &HF0
data1(694) = &H72
data1(695) = &H2
data1(696) = &H35
data1(697) = &H3
data1(702) = &HC4
data1(704) = &HFF
data1(706) = &H21
data1(707) = &H2
data1(708) = &H5A
data1(709) = &H2
data1(714) = &HC8
data1(716) = &HC0
data1(718) = &H1F
data1(719) = &H1
data1(720) = &H5F
data1(721) = &H2
data1(722) = &HC8
data1(725) = &HF0
data1(726) = &HD1
data1(728) = &HE9
data1(730) = &HCD
data1(731) = &H2
data1(732) = &HF4
data1(733) = &H3
data1(738) = &HD4
data1(740) = &HB9
data1(741) = &H1
data1(742) = &HB6
data1(743) = &H2
data1(744) = &HA3
data1(745) = &H3
data1(750) = &HD5
data1(752) = &H4
data1(753) = &H1
data1(754) = &H94
data1(755) = &H1
data1(756) = &H98
data1(757) = &H2
data1(762) = &HDA
data1(764) = &HC
data1(765) = &H2
data1(766) = &HC3
data1(767) = &H4
data1(768) = &HAD
data1(769) = &H5
data1(774) = &HE1
data1(776) = &HB4
data1(778) = &H41
data1(779) = &H1
data1(780) = &H83
data1(781) = &H1
data1(786) = &HE4
data1(788) = &H46
data1(789) = &H1
data1(790) = &H21
data1(791) = &H3
data1(792) = &HE9
data1(793) = &H4
data1(798) = &HE9
data1(800) = &H53
data1(801) = &H2
data1(802) = &HA2
data1(803) = &H2
data1(804) = &H1
data1(805) = &H4
data1(810) = &HEA
data1(812) = &H47
data1(813) = &H1
data1(814) = &HCA
data1(815) = &H1
data1(816) = &H9A
data1(817) = &H2
data1(822) = &HF7
data1(824) = &HDA
data1(825) = &H1
data1(826) = &HB2
data1(827) = &H2
data1(828) = &H9B
data1(829) = &H3
data1(834) = &H2
data1(835) = &H1
data1(836) = &H8D
data1(837) = &H2
data1(838) = &HA6
data1(839) = &H4
data1(840) = &H73
data1(841) = &H6
data1(846) = &HA
data1(847) = &H1
data1(848) = &HC2
data1(849) = &H2
data1(850) = &HE9
data1(851) = &H4
data1(852) = &HB5
data1(853) = &H6
data1(858) = &HB
data1(859) = &H1
data1(860) = &H61
data1(861) = &H2
data1(862) = &HAB
data1(863) = &H4
data1(864) = &H42
data1(865) = &H6
data1(870) = &H16
data1(871) = &H1
data1(872) = &H2B
data1(873) = &H1
data1(874) = &H28
data1(875) = &H2
data1(876) = &H17
data1(877) = &H3
data1(882) = &H1F
data1(883) = &H1
data1(884) = &H2D
data1(885) = &H1
data1(886) = &H6F
data1(887) = &H2
data1(888) = &H3
data1(889) = &H4
data1(890) = &H2C
data1(891) = &H1
data1(893) = &HF0
data1(894) = &H2C
data1(895) = &H1
data1(896) = &HDA
data1(897) = &H1
data1(898) = &H7E
data1(899) = &H3
data1(900) = &H28
data1(901) = &H5
data1(906) = &H32
data1(907) = &H1
data1(908) = &H48
data1(909) = &H1
data1(910) = &H9
data1(911) = &H3
data1(912) = &H43
data1(913) = &H5
data1(918) = &H55
data1(919) = &H1
data1(920) = &H12
data1(921) = &H1
data1(922) = &H3F
data1(923) = &H3
data1(924) = &H3D
data1(925) = &H4
data1(930) = &H70
data1(931) = &H1
data1(932) = &HD0
data1(933) = &H1
data1(934) = &H14
data1(935) = &H3
data1(936) = &H9
data1(937) = &H4
data1(942) = &H87
data1(943) = &H1
data1(944) = &HAE
data1(946) = &HAE
data1(948) = &HAE
data1(954) = &H8D
data1(955) = &H1
data1(956) = &H79
data1(957) = &H1
data1(958) = &H6A
data1(959) = &H3
data1(960) = &H8C
data1(961) = &H4
data1(966) = &H95
data1(967) = &H1
data1(968) = &H3
data1(969) = &H1
data1(970) = &HE4
data1(971) = &H1
data1(972) = &HBB
data1(973) = &H2
data1(978) = &H99
data1(979) = &H1
data1(980) = &H6E
data1(981) = &H1
data1(982) = &H9F
data1(983) = &H2
data1(984) = &HD1
data1(985) = &H3
data1(990) = &HA0
data1(991) = &H1
data1(992) = &H53
data1(993) = &H1
data1(994) = &HAC
data1(995) = &H2
data1(996) = &H1C
data1(997) = &H5
data1(1002) = &HA1
data1(1003) = &H1
data1(1004) = &H55
data1(1005) = &H1
data1(1006) = &HA3
data1(1007) = &H2
data1(1008) = &H96
data1(1009) = &H4
data1(1014) = &HA3
data1(1015) = &H1
data1(1016) = &HD
data1(1017) = &H2
data1(1018) = &H9B
data1(1019) = &H3
data1(1020) = &H63
data1(1021) = &H5
data1(1026) = &HA6
data1(1027) = &H1
data1(1028) = &H51
data1(1029) = &H1
data1(1030) = &H5D
data1(1031) = &H2
data1(1032) = &H6D
data1(1033) = &H5
data1(1038) = &HB2
data1(1039) = &H1
data1(1040) = &HAD
data1(1041) = &H1
data1(1042) = &H19
data1(1043) = &H2
data1(1044) = &HC7
data1(1045) = &H4

loadserver(0) = &H37
loadserver(1) = &H0
loadserver(2) = &HF4
loadserver(3) = &H3
loadserver(4) = &H54
loadserver(5) = &HDF
loadserver(6) = &H0
loadserver(7) = &H1E
loadserver(8) = &H4D
loadserver(9) = &H27
loadserver(10) = &H47
loadserver(11) = &HD5
loadserver(12) = &H5E
loadserver(13) = &HC7
loadserver(14) = &HCC
loadserver(15) = &H1A

loadserver(16) = &HA8
loadserver(17) = &H52
loadserver(18) = &H87
loadserver(19) = &HA9
loadserver(20) = &H80
loadserver(21) = &HFD
loadserver(22) = &HC2
loadserver(23) = &HF
loadserver(24) = &H78
loadserver(25) = &H9D
loadserver(26) = &H49
loadserver(27) = &H1A
loadserver(28) = &H2D
loadserver(29) = &H50
loadserver(30) = &HFB
loadserver(31) = &H54

loadserver(32) = &HB
loadserver(33) = &HE2
loadserver(34) = &H65
loadserver(35) = &H6B
loadserver(36) = &HC6
loadserver(37) = &H96
loadserver(38) = &HF5
loadserver(39) = &HCB
loadserver(40) = &H5
loadserver(41) = &HF5
loadserver(42) = &H89
loadserver(43) = &HD9
loadserver(44) = &HC6
loadserver(45) = &H8C
loadserver(46) = &H13
loadserver(47) = &H23

loadserver(48) = &HDC
loadserver(49) = &H4
loadserver(50) = &HD4
loadserver(51) = &HD5
loadserver(52) = &HC8
loadserver(53) = &H2F
loadserver(54) = &H33

room(0) = &H44
room(1) = &HF
room(2) = &HEB
room(3) = &H3
room(4) = &H2C
room(5) = &H1

     
ii = 12
For i = 1 To 20
            room(ii) = 120
            room(ii + 4) = &H60
            room(ii + 8) = 5
            room(ii + 11) = 5
ii = ii + 13
Next
Dim iii As Integer
ii = 398
For iii = 1 To 9
            For i = 1 To 30
                        room(ii) = iii
            ii = ii + 13
            Next
Next

wsk(0).Listen
For i = 1 To 5
            Load wsk(i)
Next
Load wskhttp(1)
wskhttp(0).Listen
Exit Sub

xXx:
Call enters
End
End Sub


Private Function resetwsk()
Dim i As Integer
For i = 1 To 5
            wsk(i).Close
Next
End Function


Private Function dosend(it As Integer, Index As Integer)
wsk(Index).SendData spass()
End Function

Private Sub wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim i As Integer
For i = 1 To 5
            If wsk(i).State <> sckConnected Then
                        wsk(i).Close
                        wsk(i).Accept requestID
                        Exit For
            End If
Next
If i = 6 Then
            add ("连接数已满!拒绝接入...")
            Exit Sub
End If
            
End Sub

Private Sub wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim abc() As Byte
Dim fun As String
Dim iover As Integer
wsk(Index).GetData abc(), vbArray + vbByte

If abc(0) + abc(1) * 256 <> bytesTotal Then
            addsock Index, "非法的数据包被发现"
End If


fun = Hex(abc(3)) & Hex(abc(2))
            
Debug.Print fun
Select Case fun
            Case "FB5"
                        addsock Index, "尝试返回房间！"
                        p9(2) = &HB6
                        p9(3) = &HF
                        p9(4) = &H0
                        p9(5) = &HE
                        p9(6) = &H0
                        p9(7) = &H0
                        p9(8) = &H0
                        wsk(Index).SendData p9()
            Case "FB0"
                        addsock Index, "一曲结束!!"
                        p6(2) = &HB1
                        p6(3) = &HF
                        p6(4) = &H0
                        p6(5) = &H1
                        wsk(Index).SendData p6()
                        For iover = 13 To 24
                                    finish(iover) = abc(iover - 9)
                        Next
                        finish(25) = abc(18)
                        finish(26) = abc(19)
                        finish(27) = abc(20)
                        finish(28) = abc(21)
                        wsk(Index).SendData finish()
                        add "Cool:" & abc(4) + abc(5) * 256 & " Good:" & abc(6) + abc(7) * 256 & " Miss:" & abc(8) + abc(9) * 256 & " Bad:" & abc(10) + abc(11) * 256 & " HI:" & abc(12) + abc(13) * 256 & " Jam:" & abc(14) + abc(15) * 256 & " Total:" & abc(18) + abc(19) * 256 + abc(20) * 65535
                        'For iover = 0 To bytesTotal - 1
                        '            Debug.Print abc(iover)
                        'Next
            Case "1771"
                        addsock Index, "仍然连接!"
                        wsk(Index).SendData abc()
            Case "FA0"
                        addsock Index, "成功登入房间！"
                        abc(2) = &HA1
                        wsk(Index).SendData abc()
            Case "FBE"
                        addsock Index, "房间获得中..."
                        wsk(Index).SendData data1()
            Case "FB7"
                        addsock Index, "服务器数据处理.."
                        ReDim jz(bytesTotal) As Byte
                        jz(0) = bytesTotal + 1
                        jz(2) = &HB8
                        jz(3) = &HF
                        jz(4) = &H1
               
                        For iover = 5 To bytesTotal
                                    jz(iover) = abc(iover - 1)
                        Next
                        wsk(Index).SendData jz()
            Case "7D0"
                        addsock Index, "加载用户信息..."
                        wsk(Index).SendData userinfo()
            Case "FA4"
                        addsock Index, "变更颜色！拒绝!"
                        p6(2) = &HA5
                        p6(3) = &HF
                        p6(4) = &H0
                        p6(5) = &H5
            Case "7D4"
                        addsock Index, "为用户建立房间."
                        wsk(Index).SendData createroom()
            Case "FB2"
                        addsock Index, "爽色你!!"
                        p4(2) = &HB5
                        p4(3) = &HF
                        wsk(Index).SendData p4()
            Case "FAA"
                        addsock Index, "开启音乐..."
                        p12(2) = &HAB
                        p12(3) = &HF
                        p12(4) = &H0
                        p12(5) = &H0
                        p12(6) = &H0
                        p12(7) = &H0
                        p12(8) = &H9C
                        p12(9) = &HED
                        p12(10) = &H7B
                        p12(11) = &HE
                        wsk(Index).SendData p12()
            Case "FBB"
                        addsock Index, "退出房间.."
                        abc(2) = &HBC
                        wsk(Index).SendData abc()
            Case "FA2"
                        addsock Index, "FA2 Request"
                        p6(2) = &HA5
                        p6(3) = &HF
                        p6(4) = &H0
                        p6(5) = &H0
                        wsk(Index).SendData p6()
            Case "FAC"
                        addsock Index, "音乐就绪.."
                        p5(2) = &HAD
                        p5(3) = &HF
                        p5(4) = &H0
                        wsk(Index).SendData p5()
            Case "FAA"
                        addsock Index, "嘎快就死外的?"
                        p12(2) = &HAB
                        p12(3) = &HF
                        p12(4) = &H0
                        p12(5) = &H0
                        p12(6) = &H0
                        p12(7) = &H0
                        p12(8) = &H9B
                        p12(9) = &H50
                        p12(10) = &HB7
                        p12(11) = &H2
                        wsk(Index).SendData p12()
            Case "BBD"
                        addsock Index, "退出房间.."
                        p8(2) = &HBE
                        p8(3) = &HB
                        p8(4) = &H0
                        p8(5) = &H0
                        p8(6) = &H0
                        p8(7) = &H0
                        wsk(Index).SendData p8()
            Case "BB8"
                        addsock Index, "更改房间名称！拒绝"
                        p10(2) = &HB9
                        p10(3) = &HB
                        p10(4) = Asc("%")
                        p10(5) = Asc("H")
                        p10(6) = Asc("O")
                        p10(7) = Asc("H")
                        p10(8) = Asc("O")
                        p10(9) = Asc("%")
                        wsk(Index).SendData p10()
            Case "FAC"
                        addsock Index, "准备歌曲..."
                        p5(2) = &HAD
                        p5(3) = &HF
                        p5(4) = &H0
                        wsk(Index).SendData p5()
            Case "7E8"
                        p12(2) = &HE7
                        p12(3) = &H7
                        p12(4) = &H0
                        p12(5) = &H0
                        p12(6) = &H0
                        p12(7) = &H0
                        p12(8) = &HB2
                        p12(9) = &H0
                        p12(10) = &H0
                        p12(11) = &H1
                        wsk(Index).SendData p12()
            Case "7E5"
                        addsock Index, "返回星球"
                        p8(2) = &HE6
                        p8(3) = &H7
                        p8(4) = &H0
                        p8(5) = &H0
                        p8(6) = &H0
                        p8(7) = &H0
                        wsk(Index).SendData p8
            Case "3EA"
                        addsock Index, "开始获取星球信息！"
                        wsk(Index).SendData room()
            Case "3F3"
                        addsock Index, "要求获得星球信息!"
                        wsk(Index).SendData loadserver()
            Case "FB9"
                        addsock Index, "专辑房间.."
                        p6(2) = &HBA
                        p6(3) = &HF
                        p6(4) = &H0
                        p6(5) = &H0
                        wsk(Index).SendData p6()
            Case "3F4", "3E8"
                        addsock Index, "准备获得星球信息!"
                        p8(2) = &HE9
                        p8(3) = &H3
                        p8(4) = &H0
                        p8(5) = &H0
                        p8(6) = &H0
                        p8(7) = &H0
                        wsk(Index).SendData p8()
            Case "13A4"
                        addsock Index, "分配客户空间!"
                        p8(2) = &HA5
                        p8(3) = &H13
                        p8(4) = &H91
                        p8(5) = &H2E
                        p8(6) = &H0
                        p8(7) = &H0
                        wsk(Index).SendData p8
            Case "3F1"
                        addsock Index, "尝试登陆游戏"
                        wsk(Index).SendData spass()
            Case "7D2"
                        addsock Index, "获得房间信息"
                        wsk(Index).SendData data2()
                        wsk(Index).SendData data3()
            Case "3EC"
                        addsock Index, "正在进入房间"
                        p16(2) = &HED
                        p16(3) = &H3
                        p16(4) = &H0
                        p16(5) = &H0
                        p16(6) = &H0
                        p16(7) = &H0
                        p16(8) = &H0
                        p16(9) = &H0
                        p16(10) = &H0
                        p16(11) = &H0
                        p16(12) = &HE4 'A0
                        p16(13) = &H8 '6
                        p16(14) = &H0
                        p16(15) = &H0
                        wsk(Index).SendData p16()
            Case "3EF"
                        If abc(0) = 13 Then
                                    Dim sum As Long
                                    sum = abc(4) * 13 + abc(6) * 14 + abc(7) * 15 + abc(8) * 16 + abc(9) * 17 + abc(10) * 18 + abc(11) * 19 + abc(5)
                                    Debug.Print "sum" & sum
                                    If sum = 12345 Then
                                                addsock Index, "登陆了游戏!"
                                                p12(2) = &HF0
                                                p12(3) = &H3
                                                p12(4) = &H0
                                                p12(5) = &H0
                                                p12(6) = &H0
                                                p12(7) = &H0
                                                p12(8) = &H90
                                                p12(9) = &H23
                                                p12(10) = &H1
                                                p12(11) = &H0
                                                wsk(Index).SendData p12()
                                                Exit Sub
                                    End If
                        End If
                        addsock Index, "登陆失败!"
                        p8(2) = &HF0
                        p8(3) = &H3
                        p8(4) = &HFF
                        p8(5) = &HFF
                        p8(6) = &HFF
                        p8(7) = &HFF
                        wsk(Index).SendData p8()
            Case "FFF0"
                        addsock Index, "客户主动要求断开一个连接"
                        wsk(Index).Close
            Case Else
                        addsock Index, "未知命令:" & fun
            
End Select
If Err Then addsock Index, "接收数据出现一个错误!连接已关闭!"
End Sub


Private Sub wsk_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
add ("套接字错误id" & Index)
wsk(Index).Close
End Sub


Private Function add(str As String)
Text1.Text = Time & ": " & str & vbCrLf & Text1.Text
End Function

Private Function addsock(Index As Integer, str As String)
Text1.Text = Time & "," & wsk(Index).RemoteHostIP & ": " & str & vbCrLf & Text1.Text
End Function

Private Sub wskhttp_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Debug.Print "okkok"
wskhttp(1).Close
wskhttp(1).Accept requestID
End Sub

Private Sub wskhttp_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim str As String
wskhttp(Index).GetData str
If InStr(1, str, "txt") Then
            wskhttp(Index).SendData httpheader & "599" & vbCrLf & vbCrLf & httpmsg1 & httpmsg2
ElseIf InStr(1, str, "asp") Then
            wskhttp(Index).SendData httpheader & "192" & vbCrLf & vbCrLf & "<html><body bgcolor=""#0095BF"" leftmargin=""0"" topmargin=""0"" oncontextmenu=""return false"" onselectstart=""return false"" ondragstart=""return false"">O2Jam Personal Server By HOHO</html>"
End If
End Sub

Private Sub wskhttp_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskhttp(Index).Close
End Sub

Private Sub wskhttp_SendComplete(Index As Integer)
wskhttp(1).Close
End Sub
