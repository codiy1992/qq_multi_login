Attribute VB_Name = "GlobalData"
Option Explicit
'====================================================ȫ��������========================================================
Public Yzmcode         As String ' Dialog  ���ڱ�����֤��
Public Pdunload        As Boolean 'Dialog  �����Ƿ�����
Public Codeqq          As Long '   Dialog  ���ڶ�ȡ��֤��QQ����
'----------------------------------------------------------------
Public Const MAX_NUM = 25
Public ID              As Long     '��½ID
Public cuTotalCount As Long
Public canLoginNew     As Boolean  '�Ƿ��ڵ�½
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
