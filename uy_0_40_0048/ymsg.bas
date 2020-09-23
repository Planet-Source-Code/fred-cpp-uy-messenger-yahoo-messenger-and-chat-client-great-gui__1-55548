Attribute VB_Name = "ymsg"
Option Explicit

'Public options Variables.
Public bAceptIMs As Integer         'Implemented
Public bAutoJoinRoom As Boolean
Public bAutoReConnect As Boolean
Public bShowSmileys As Boolean
Public bRememberID As Boolean
Public bRememberPass As Boolean
Public sLastID As String
Public sDefaultRoom As String

Public bChatEmmbeded As Boolean     'wip
Public sProtocol As String          'wip
Public bUseCustomFonts As Boolean   'wip

Public Sub SaveSettings()
    'Save the global Settings
    SaveSetting App.Title, "Settings", "bRememberID", CStr(bRememberID)
    SaveSetting App.Title, "Settings", "bAceptIMs", CStr(bAceptIMs)
    SaveSetting App.Title, "Settings", "bShowSmileys", CStr(bShowSmileys)
    SaveSetting App.Title, "Settings", "bAutoJoinRoom", CStr(bAutoJoinRoom)
    SaveSetting App.Title, "Settings", "bAutoReConnect", CStr(bAutoReConnect)
    SaveSetting App.Title, "Settings", "bRememberPass", CStr(bRememberPass)
    SaveSetting App.Title, "Settings", "sLastID", CStr(sLastID)
    SaveSetting App.Title, "Settings", "sDefaultRoom", CStr(sDefaultRoom)
End Sub

Public Sub ReadSettings()
    'Read the global settings
    bRememberID = CBool(GetSetting(App.Title, "Settings", "RememberID", False))
    bAceptIMs = CBool(GetSetting(App.Title, "Settings", "bAceptIMs", "1"))
    bShowSmileys = CBool(GetSetting(App.Title, "Settings", "bShowSmileys", "1"))
    bAutoJoinRoom = CBool(GetSetting(App.Title, "Settings", "bAutoJoinRoom", "1"))
    bAutoReConnect = CBool(GetSetting(App.Title, "Settings", "bAutoReConnect", "1"))
    bRememberPass = CBool(GetSetting(App.Title, "Settings", "bRememberPass", "1"))
    sLastID = GetSetting(App.Title, "Settings", "sLastID", "")
    sDefaultRoom = GetSetting(App.Title, "Settings", "sDefaultRoom", "")
End Sub

Public Sub MPRecibed(data As String)
    'YMSG*****~****·pC 5À€module_ocxÀ€4À€freddemilenaÀ€97À€1À€14À€<FADE #011e83,#022ec4,#011e83>[1m<font face="Verdana">Hola</FADE>À€63À€;0À€64À€0À€
    Dim ni As Integer, iTmp As Integer
    Dim sMsgTo As String, sMsgFrom As String, sMessage As String
    Dim tmp
    If bAceptIMs = 0 Then Exit Sub   'Dont waste time and resources if not acept's PM's
    'GetTo
    On Error GoTo unknownPacket
    ni = InStr(24, data, "À€", vbTextCompare)
    sMsgTo = Mid(data, 24, ni - 24)  ' InStr(ni + 1, data, "À€", vbTextCompare) - ni - 2)
    'GetFrom
    ni = InStr(ni + 2, data, "À€", vbTextCompare)
    sMsgFrom = Mid(data, ni + 2, InStr(ni + 1, data, "À€", vbTextCompare) - ni - 2)
    'GetMessage
    ni = InStr(ni + 2, data, "14À€", vbTextCompare)
    'Debug.Print data
    sMessage = Mid(data, ni + 4, InStr(ni + 4, data, "À€", vbTextCompare) - ni - 4)
    If bAceptIMs = 1 Then
    'Acept Instant Messages from Friends Only
        On Error GoTo NoFriend
        iTmp = frmMain.tvFriends.Nodes(sMsgFrom).Index
        If iTmp <> 0 Then
            ShowIM sMsgTo, sMsgFrom, sMessage
            Exit Sub
        End If
NoFriend:
        On Error GoTo nocontacted
        For Each tmp In Forms
            If tmp.Tag = sMsgFrom Then
                ShowIM sMsgTo, sMsgFrom, sMessage
                Exit Sub
            End If
        Next
    Else
    'Acept IM from everyone
        ShowIM sMsgTo, sMsgFrom, sMessage
    End If
    ''Debug.Print err.Description
    'Show Alert
    Exit Sub
nocontacted:
    Exit Sub
unknownPacket:
    'Debug.Print "Unknown packet, maybe you are under attrack"
End Sub

Public Sub ShowIM(sTo As String, sFrom As String, sMessage As String)
    ' Add the IM on the proper window
    'Open or Activate a pm window o a especified user.
    Dim pm As Form
    For Each pm In Forms
        If pm.Tag = sFrom Then
            pm.SetFocus
            pm.AddMessage sMessage
            pm.txtTo.Text = sFrom
            Exit Sub
        End If
    Next
    'Exit Sub
    Set pm = New frmPM
    pm.Tag = sFrom
    pm.Caption = sTo & " - " & sFrom
    pm.txtTo.Text = sFrom
    pm.Show
    pm.AddMessage sMessage
    pm.Flash
    
End Sub

Public Sub TPMsg(data As String)
    Dim ni As Integer, iTmp As Integer
    Dim sMsgTo As String, sMsgFrom As String, sMessage As String
    Dim tmp
    If bAceptIMs = 0 Then Exit Sub   'Dont waste time and resources if not acept's PM's
    'GetTo
    On Error GoTo unknownPacket
    ni = InStr(24, data, "À€", vbTextCompare)
    sMsgTo = Mid(data, 24, ni - 24)
    'GetFrom
    ni = InStr(ni + 2, data, "À€", vbTextCompare)
    sMsgFrom = Mid(data, ni + 2, InStr(ni + 1, data, "À€", vbTextCompare) - ni - 2)
    If bAceptIMs = 1 Then
        'Notify only from users I have an Open PM Window
        On Error GoTo nocontacted
        For Each tmp In Forms
            If tmp.Tag = sMsgTo Then 'was msgfrom
                NotifyTyping sMsgTo, sMsgFrom
                Exit Sub
            End If
        Next
    End If
nocontacted:
unknownPacket:
    Exit Sub
End Sub


Public Sub NotifyTyping(sTo As String, sFrom As String)
    ' If the user Is typing is in the user list that can contact us
    'notify
    Dim pm As Form
    For Each pm In Forms
        If pm.Tag = sTo Then
            pm.Typing
            Exit Sub
        End If
    Next
End Sub
