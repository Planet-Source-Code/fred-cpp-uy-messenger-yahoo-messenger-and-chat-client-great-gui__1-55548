Attribute VB_Name = "modyconn"
Private Declare Function GetYahooStrings Lib "YM11AUTH.DLL" (ByVal username As String, ByVal PassWord As String, ByVal seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean

'all code by: shinkaiho if you use anything give me credit
Public Const Host_Name As String = "cs43.msg.dcn.yahoo.com"
Const name As String = "YMSG"
Const Ver As Integer = 11

Public Function Header(ByVal PacketType As String, ByVal Packet As String) As String
Header = name & Chr$(0) & Chr(Ver) & String(3, 0) & Chr(Len(Packet)) & Chr(0) & _
Chr("&H" & PacketType) & String(8, 0) & Packet
''Debug.Print Header
End Function

Public Function Get_Key(ByVal username As String) As String
Get_Key = "1À€" & username & "À€"
Get_Key = Header(57, Get_Key)
End Function

Public Function Login(ByVal username As String, ByVal crypt1 As String, ByVal crypt2 As String) As String
    Login = "6À€" & crypt1 & "À€96À€" & crypt2 & "À€0À€" & username & "À€2À€1À€1À€" & username & _
    "À€135À€5, 6, 0, 1347À€148À€300À€"
    Login = Header(54, Login)
End Function

Function SendIM(ByVal Sender, ByVal Recipient As String, ByVal Message As String)

    Dim Packet1 As String
    Dim Packet2 As String

    On Error GoTo Noymsgaviable
    Packet1 = "1À€" & Sender & "À€5À€" & Recipient & "À€14À€" & Message & "À€97À€1À€63À€;0À€64À€0À€"
    Packet2 = Header("6", Packet1)
    
    frmMain.wsYahoo.SendData Packet2
    
Noymsgaviable:
    If err.Number = 40006 Then frmMain.YCht.ychtSendPrivateMessage Recipient, Message

''Debug.Print Packet2
End Function


'Public Function getencrstrings(name As String, pass As String, seed As String, str1 As String, str2 As String, mode As Long) As Boolean
'
'Dim s1 As String, s2 As String, n As Long
'    On Error GoTo err
'    s1 = String(100, vbNullChar)
'    s2 = String(100, vbNullChar)
'    getencrstrings = GetYahooStrings(name, pass, seed, s1, s2, mode)
'    n = InStr(1, s1, vbNullChar)
'    str1 = Left$(s1, n - 1)
'    n = InStr(1, s2, vbNullChar)
'    str2 = Left$(s2, n - 1)
'    Exit Function
'err:
'    getencrstrings = False
'End Function
'


