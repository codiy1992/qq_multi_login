Attribute VB_Name = "全局Function"
Option Explicit
Public Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String, lpdwSize As Long) As Boolean
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const CP_UTF8 = 65001
'-----------------结合类模块"clswaitabletimer"使用-------------
Public mobjWaitTimer As clswaitabletimer


Function GetRnd(ByVal n As Integer) As String '获取N位随机数
Randomize
Const Cstring As String = "1234567890"
GetRnd = Mid("0" & Rnd(1) & Cstring, 1, n)
End Function
Function GetRnd1(ByVal n As Integer)  '获取N位随机数
Dim X
Randomize
GetRnd1 = Int(8 * Rnd + 1)
For X = 2 To n
    Randomize
    GetRnd1 = GetRnd1 & Int(9 * Rnd + 0)
Next
End Function
Function Encode(P As String, Key As String, code As String)
    Dim Pass As String, Jsc As String
    Jsc = Jsc & "function getp(){"
    Jsc = Jsc & "var I=hexchar2bin(md5(""" & Trim(P) & """));" '密码
    Jsc = Jsc & "var H=md5(I+""" & Key & """);" 'KEY
    Jsc = Jsc & "var G=md5(H+""" & Trim(UCase(code)) & """);" '验证码
    Jsc = Jsc & "return G;}"
    Form1.ScriptControl1.AddCode Jsc
    Encode = Form1.ScriptControl1.Run("getp") '加密密码
End Function
Function get_gtk(sk_ey As String)
    Dim js(6) As String
    js(0) = "function getGTK(str){" & vbCrLf
    js(1) = "var hash = 5381;" & vbCrLf
    js(2) = "for(var i = 0, len = str.length; i < len; ++i){" & vbCrLf
    js(3) = "    hash += (hash << 5) + str.charAt(i).charCodeAt();" & vbCrLf
    js(4) = "}" & vbCrLf
    js(5) = " return hash & 0x7fffffff;" & vbCrLf
    js(6) = "}"
  
    Form1.ScriptControl1.AddCode js(0) & js(1) & js(2) & js(3) & js(4) & js(5) & js(6)
    get_gtk = Form1.ScriptControl1.Run("getGTK", sk_ey)
End Function

'inet请求获取数据有时会出现乱码，需要用此函数转码
Function BytesToUnicode(ByRef Utf() As Byte) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    lLength = UBound(Utf) - LBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    BytesToUnicode = String$(lBufferSize, Chr(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(BytesToUnicode), lBufferSize)
    If lRet <> 0 Then
        BytesToUnicode = Left(BytesToUnicode, lRet)
    Else
        BytesToUnicode = ""
    End If
End Function
Function Unmid(StrU, Minstr, Maxstr) '取中间文本
'If InStr(StrU, Minstr) > 0 And InStr(StrU, Maxstr) > 0 Then
   Dim q1 As Long, q2 As Long
   q1 = InStr(StrU, Minstr) + Len(Minstr)
   q2 = InStr(q1, StrU, Maxstr)
   'Debug.Print "q2=" & q2
   If q2 = 0 Then Unmid = Replace(StrU, Left(StrU, q1), ""): Exit Function
   Unmid = Mid(StrU, q1, q2 - q1)
'Else
'   Unmid = 0
'End If
End Function
Function GetTimerc() '取时间戳
Dim cs As Date, xs As Date, t As Long
cs = CDate(Now)
xs = CDate("1970-01-01 08:00:00")
Randomize
GetTimerc = DateDiff("s", xs, cs) * 1000 + Int(1 * Rnd + 999)
End Function



'--------------------------------------------------------------
'**************结合类模块"clswaitabletimer"使用'***************
'函数：vb_Sleep 【使用内核对象WaitableTimer实现】
'参数：(时间) 单位为：毫秒
'返回值：无返回值
'--------------------------------------------------------------
Public Function vb_Sleep(dwMilliseconds As Long)
    Set mobjWaitTimer = New clswaitabletimer
            mobjWaitTimer.Wait (dwMilliseconds)
    Set mobjWaitTimer = Nothing
End Function

'Function outPutStr(ByVal str As String, Optional ByVal path As String = "log.txt")
'    Open path For Append As #1
'    Print #1, , str
'    Close #1
'End Function

Public Function Encript(ByVal strValue As String) As String
    Dim byteValue() As Byte
    Dim i As Integer
    byteValue = StrConv(strValue, vbFromUnicode)  '字符串转成数组
    For i = 0 To UBound(byteValue)
    byteValue(i) = (byteValue(i) + 5) Mod 255
    Next
    Encript = BytesToUnicode(byteValue())
End Function
Public Function Decript(ByVal strValue As String) As String
    Dim byteValue() As Byte
    Dim i As Integer
    byteValue = StrConv(strValue, vbFromUnicode)  '字符串转成数组
    For i = 0 To UBound(byteValue)
    byteValue(i) = (byteValue(i) - 5) Mod 255
    Next
    Decript = BytesToUnicode(byteValue())
End Function

