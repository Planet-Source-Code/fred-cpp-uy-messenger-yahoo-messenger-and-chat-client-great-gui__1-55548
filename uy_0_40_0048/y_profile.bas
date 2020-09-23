Attribute VB_Name = "y_profile"
Option Explicit

''''''''''''''''''''''''''''''''
' Profiler 2.0
' ------------
'
' Yahoo_Profile.bas - Michael Bacoz
' http://www.Fatal-Instinct.com
''''''''''''''''''''''''''''''''
' Edited to be useful in this project
'   by Fred.cpp

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC

Public LastImage As String

Public Function Age(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Age:</font>")
If Separator Then
    First = Mid$(strText, Separator + 53)
    Age = Left$(First, InStr(First, "</b>") - 1)
End If
If Len(Age) <= 1 Then Age$ = "Not Specified"
End Function

Public Function Gender(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Gender:</font>")
If Separator Then
    First = Mid$(strText, Separator + 56)
    Gender = Left$(First, InStr(First, "</b>") - 1)
End If

If Len(Gender) <= 1 Then Gender$ = "Not Specified"
End Function

Public Function Last_Updated(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, " <b>")
If Separator Then
    First = Mid$(strText, Separator + 4)
    Last_Updated = Left$(First, InStr(First, "</font>") - 5)
End If
End Function

Public Function Location(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Location:</font>")
If Separator Then
    First = Mid$(strText, Separator + 57)
    Location = Left$(First, InStr(First, "</b>") - 1)
End If

If Len(Location) <= 1 Then Location$ = "Not Specified"
End Function

Public Function Marital_Status(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Marital&nbsp;Status:</font>")
If Separator Then
    First = Mid$(strText, Separator + 69)
    Marital_Status = Left$(First, InStr(First, "</b>") - 1)
End If

If Len(Marital_Status) <= 1 Then Marital_Status$ = "Not Specified"
End Function

Public Function NickName(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Nickname:</font></td>")
If Separator Then
    First = Mid$(strText, Separator + 56)
    NickName = Left$(First, InStr(First, "</b>") - 1)
End If

If Len(NickName) <= 1 Then NickName$ = "Not Specified"
End Function

Public Function Occupation(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Occupation:</font>")
If Separator Then
    First = Mid$(strText, Separator + 59)
    Occupation = Left$(First, InStr(First, "</b>") - 1)
End If

If Len(Occupation) <= 1 Then Occupation$ = "Not Specified"
End Function
Public Function OnOff(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "ltn.gif")
If Separator Then
    First = Mid$(strText, Separator + 86)
    OnOff = Left$(First, InStr(First, "</b") - 2) ' - 105)
End If

If Len(OnOff) <= 1 Then OnOff$ = "I'm Offline"
End Function
Public Function RealName(strText As String) As String
Dim Separator As Integer
Dim First As String
Separator = InStr(strText, "Real&nbsp;Name:</font></td>")
If Separator Then
    First = Mid$(strText, Separator + 62)
    RealName = Left$(First, InStr(First, "</b>") - 1)
End If

If Len(RealName) <= 1 Then RealName$ = "Not Specified"
End Function

Public Function Send_Text(strText As String)
Dim imclass As Long, Button As Long, richedit As Long

imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, strText$)

Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function Strip_URL(FileName) As String
Dim NewFilename As String

NewFilename$ = Mid(FileName, InStrRev(FileName, "/") + 1)
Strip_URL = Left$(NewFilename$, InStrRev(NewFilename$, "?") - 1)
End Function

Public Function Yahoo_ID(strText As String) As String
Dim Separator As Integer
Dim First As String
Dim lngSecond As String
Separator = InStr(strText, "Yahoo! ID:</font>")
If Separator Then
    First = Mid$(strText, Separator + 60)
    lngSecond = Left$(First, InStr(First, "</b>") - 1)
    Yahoo_ID = Mid$(lngSecond, InStr(lngSecond, "<b>") + 3)
End If
End Function



Public Function GrabUserData(inetObj As Inet, strUsername As String)
    'Get the Full Data for the user.
    Dim FileBuffer() As Byte
    On Error Resume Next
    Dim Separator As Integer
    Dim before As String
    Dim URL As String
    Dim FileName As String
    Dim strHTML As String
    
    inetObj.Cancel
    'Main.Inet2.Cancel
    'Command2.Enabled = False
    If strUsername = "" Then GrabUserData = ""
    
    If inetObj.StillExecuting Then DoEvents
    strHTML$ = inetObj.OpenURL("http://profiles.yahoo.com/" & strUsername)
    If inetObj.StillExecuting Then DoEvents
        
    If InStr(1, UCase(strHTML$), UCase("Sorry, but the profile")) Then
        GrabUserData = "The Selected User Doesn't Exist"
        LastImage = ""
        Exit Function
    End If
    
    If InStr(1, UCase(strHTML$), UCase("Profile contains possible")) Then
        GrabUserData = "The selected user has an Adult profile"
        LastImage = ""
        Exit Function
    End If

Separator = InStr(strHTML$, "<table cellspacing=1 cellpadding=2 border=0 width=" & Chr(34) & "1%" & Chr(34) & ">")
    
    If Separator = False Then
        Exit Function
        LastImage = ""
    End If
    before = Mid$(strHTML, Separator + 1)
    before = Mid$(before, InStr(before, "<a href=") + 9)
    URL = Left$(before, InStr(before, "<img") - 3)
    
    FileName = Strip_URL(URL)
    
    FileBuffer() = inetObj.OpenURL(URL, 1)
    
    Open App.Path & "\tmp\" & FileName$ For Binary Access Write As #1
    Put #1, , FileBuffer()
    Close #1
    
    
    'Debug.Print App.Path & "\tmp\" & FileName$
    LastImage = App.Path & "\tmp\" & FileName$
    'Pic.Picture1.Picture = LoadPicture(App.Path & "\" & Filename$)
    'pb.Value = 95
    'Kill App.Path & "\" & Filename$
    GrabUserData = "Last update: " & Last_Updated(strHTML$) & vbCrLf & _
            "Real name: " & RealName(strHTML$) & vbCrLf & _
            "Nick name: " & NickName(strHTML$) & vbCrLf & _
            "Location: " & Location(strHTML) & vbCrLf & _
            "Age: " & Age(strHTML) & vbCrLf & _
            "Marital Status: " & Marital_Status(strHTML) & vbCrLf & _
            "Gender: " & Gender(strHTML) & vbCrLf & _
            "Ocupation: " & Occupation(strHTML) & vbCrLf & _
            "Current Status: " & OnOff(strHTML)
End Function

Public Sub GetAsyncProfile(sUser As String)
    DoEvents
    frmMain.isebMain.SetDetailsText y_profile.GrabUserData(frmMain.inetGrab, sUser)
    DoEvents
    If LCase(Right(LastImage, 3)) = "jpg" Or LCase(Right(LastImage, 3)) = "gif" Then
        frmMain.isebMain.AddItem -1, "Picture", "View " & sUser & " photo", 12
        frmMain.isebMain.SetDetailsImage LoadPicture(y_profile.LastImage)
    End If
End Sub


