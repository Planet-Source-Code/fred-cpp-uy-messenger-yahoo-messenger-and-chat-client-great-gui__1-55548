Attribute VB_Name = "yahoorelated"
Option Explicit

Public sTmpFriends As String
Public sTmpIgnored As String
Public sTmpAlias As String
Public sAlias() As String

'Public wndPM() As frmPM

Public Sub GetYahooFriendsList(data As String)
    '   This routine will be called while yahoo
    '   sends info about the lisf for this account.
    '   at end, this function wll call
    '   frmMain.BuildLists()
    Dim lStart As Integer, lEnd As Integer
    If Left$(data, 4) = "YMSG" Then
        'Contains Friends list?
        lStart = InStr(19, data, "87¿Ä", vbTextCompare)
        If lStart <> -1 Then      'Friends List
            lEnd = InStr(lStart + 4, data, "¿Ä", vbTextCompare)
            sTmpFriends = sTmpFriends & Mid(data, lStart + 4, lEnd - lStart)
            If lEnd = Len(data) Then Exit Sub
        End If
        'Contains Ignored list?
        lStart = InStr(19, data, "88¿Ä", vbTextCompare)
        If lStart <> -1 Then      'Ignore List
            lEnd = InStr(lStart + 4, data, "¿Ä", vbTextCompare)
            sTmpIgnored = sTmpIgnored & Mid(data, lStart + 4, lEnd - lStart)
            If lEnd = Len(data) Then Exit Sub
        End If
        'Contains Aliases?
        lStart = InStr(19, data, "89¿Ä", vbTextCompare)
        If lStart > 1 - 1 Then    'Ignore List
            lEnd = InStr(lStart + 4, data, "¿Ä59", vbTextCompare)
            sTmpAlias = Mid(data, lStart + 4, lEnd - lStart - 4)
            frmMain.BuildLists
        End If
    End If
End Sub

Public Sub SetFriendStatus(sdata, ByRef tvtree As TreeView)
    ' This Routine will set the users Status afther have logged In
    Dim sTmp, ni As Integer
    FriendStatus Right$(sdata, Len(sdata) - 23), tvtree
End Sub

Public Sub GetFriendsStatus(sdata As String, ByRef tvtree As TreeView)
    ' This Routine will set the users Status.
    ' YMSG    5     ∑~È0¿Äfred_cpp¿Ä1¿Ä
    ' YMSG    Q     ∑z?Î0¿Äfred_cpp¿Ä1¿Ä
    Dim sUserCodes() As String
    Dim sTmp
    sUserCodes = Split(sdata, "¿Ä7¿Ä", , vbTextCompare)
    For Each sTmp In sUserCodes
        'Debug.Print sTmp
        'GetBetween stmp,
        FriendStatus sTmp, tvtree
    Next
End Sub

Public Sub SetFriendOffline(sdata As String, ByRef tvtree As TreeView)
    'Set the status of a user to offline
    ' We Already know this is a offline message,
    ' We just need to get the user name
    Dim sUser As String
    On Error GoTo analizeerr
    sUser = Mid(sdata, 24, InStr(25, sdata, "¿Ä10¿Ä", vbTextCompare) - 24)
    tvtree.Nodes.Item(sUser).Text = sUser
    tvtree.Nodes.Item(sUser).Image = 4
    tvtree.Nodes.Item(sUser).SelectedImage = 4
    'Add status notification in IM window If aviable
        Dim pm As Form
    For Each pm In Forms
        If pm.Tag = sUser Then
            pm.AddStatusNotification 4, "Offline"
            pm.Show
            Exit Sub
        End If
    Next
    'error description
    Exit Sub
analizeerr:
    If err.Number = 5 Then
        frmMain.Disconect
    End If
End Sub
Private Function FriendStatus(ByVal data As String, ByRef tvtree As TreeView) As String
    On Error Resume Next
    Dim spl() As String, spl1() As String
    Dim ni As Integer, sUser As String, code As String
    Dim iImage As Integer
    ni = InStr(1, data, "¿Ä10¿Ä", vbTextCompare)
    sUser = Left(data, ni - 1)
    code = Mid(data, ni + 6, InStr(ni + 6, data, "¿Ä", vbTextCompare) - ni - 6)
    Select Case code
        Case Is = "0"
            FriendStatus = "I'm Available"
            iImage = 0
        Case Is = "1"
            FriendStatus = "Be Right Back"
            iImage = 1
        Case Is = "2"
            FriendStatus = "Busy"
            iImage = 1
        Case Is = "3"
            FriendStatus = "Not at home"
            iImage = 1
        Case Is = "4"
            FriendStatus = "Not at my desk"
            iImage = 1
        Case Is = "5"
            FriendStatus = "Not in the office"
            iImage = 1
        Case Is = "6"
            FriendStatus = "On the phone"
            iImage = 1
        Case Is = "7"
            FriendStatus = "On Vacation"
            iImage = 1
        Case Is = "8"
            FriendStatus = "Out of lunch"
            iImage = 1
        Case Is = "9"
            FriendStatus = "Stepped Out"
            iImage = 1
        Case Is = "99"
            ni = InStr(data, "¿Ä19¿Ä")
            If ni <> -1 Then
                spl = Split(data, "¿Ä19¿Ä")
                spl1 = Split(spl(1), "¿Ä")
                FriendStatus = spl1(0)
                ni = InStr(1, data, "¿Ä47¿Ä", vbTextCompare)
                iImage = CInt(Mid(data, ni + 6, 1))
            End If
        Case Is = "999"
            FriendStatus = "Idle"
            iImage = 2
        Case Else
            FriendStatus = "Unknown Status"
            iImage = 3
    End Select
    'Set To Boldø?:/
    tvtree.Nodes.Item(sUser).Text = sUser & IIf((code = "0"), "", "(" & FriendStatus & ")")
    tvtree.Nodes.Item(sUser).Image = iImage + 5
    tvtree.Nodes.Item(sUser).SelectedImage = iImage + 5
    Dim pm As Form
    For Each pm In Forms
        If pm.Tag = sUser Then
            pm.AddStatusNotification iImage + 5, AddSmileys(FriendStatus)
            pm.Show
            Exit Function
        End If
    Next
End Function

Public Function GetBetween(IStringStr As String, IBefore As String, IPast As String)
    ' Function imported from Andy at
    ' http://venky.proboards10.com/index.cgi?action=viewprofile&username=Andy"
    
    On Error Resume Next
    Dim iString As String
    iString = IStringStr
    iString = Right(iString, Len(iString) - InStr(iString, IBefore) - Len(IBefore) + 1)
    iString = Mid(iString, 1, InStr(iString, IPast) - 1)
    GetBetween = iString
End Function

Public Sub PMUser(sUser As String)
    'Open or Activate a pm window o a especified user.
    Dim pm As Form
    For Each pm In Forms
        If pm.Tag = sUser Then
            pm.SetFocus
            Exit Sub
        End If
    Next
    Set pm = New frmPM
    pm.Tag = sUser
    pm.Caption = "Private message - " & sUser
    pm.txtTo.Text = sUser
    pm.Show
    pm.SetFocus
End Sub
