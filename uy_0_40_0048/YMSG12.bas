Attribute VB_Name = "ymsg12"
'******************************************'
'  YMSG12 Login by Markus ( www.romanware.com)
'  thanks for sharing!
Option Explicit
Private Declare Function Ymsg12Crypt Lib "YMSG12Crypt.DLL" (ByVal PassWord As String, ByVal seed As String, ByVal result6 As String, ByVal result96 As String) As Long

Public Function getencrstrings(pass As String _
                , seed As String _
                , str1 As String _
                , str2 As String _
                ) As Boolean
    Dim s1 As String, s2 As String, n As Long
    On Error GoTo err
    s1 = String(80, vbNullChar)
    s2 = String(80, vbNullChar)
    getencrstrings = Ymsg12Crypt(pass, seed, s1, s2)
    n = InStr(1, s1, vbNullChar)
    str1 = Left$(s1, n - 1)
    n = InStr(1, s2, vbNullChar)
    str2 = Left$(s2, n - 1)
    Exit Function
err:
    getencrstrings = False
End Function



