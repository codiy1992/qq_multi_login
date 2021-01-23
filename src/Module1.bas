Attribute VB_Name = "GlobalData"
Option Explicit
'====================================================全局数据区========================================================
Public Yzmcode         As String ' Dialog  窗口保存验证码
Public Pdunload        As Boolean 'Dialog  窗口是否销毁
Public Codeqq          As Long '   Dialog  窗口读取验证码QQ号码
'----------------------------------------------------------------
Public Const MAX_NUM = 25
Public ID              As Long     '登陆ID
Public cuTotalCount As Long
Public canLoginNew     As Boolean  '是否在登陆
Public status            As String
Public User(1 To MAX_NUM) As String
Public Pass(1 To MAX_NUM) As String
Public sKey(1 To MAX_NUM) As String
Public clientID(1 To MAX_NUM) As String
Public vfWebQQ(1 To MAX_NUM) As String
Public sessionID(1 To MAX_NUM) As String


Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
'Public hInternetSession As Long
'Dim hInternet As Long
Public lpTrayIconData As NOTIFYICONDATA
Public showTray As Boolean
