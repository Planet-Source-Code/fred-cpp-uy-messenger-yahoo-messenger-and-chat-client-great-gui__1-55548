Attribute VB_Name = "y_smileys"
Option Explicit

'I use this string to have the smileys definitions
'aviable from every window I create in the program.
Public strSmiley() As String

'Load a file into a string array
Public Function LoadArrayFile(ByRef FileName As String) As String()
    Dim f As Integer, xFN As String
    Dim ITries As Long
    Dim FileString As String
    On Error GoTo LoadArrayFile_ErrHandler
    f = FreeFile
    Open FileName For Binary Access Read As #f
    FileString = Space(LOF(f))
    Get f, , FileString
    LoadArrayFile = Split(FileString, Chr$(10))
    Close #f
    Exit Function
LoadArrayFile_ErrHandler:
End Function

'Load the smileys equivalents definition file into the array
Public Sub LoadSmileys()
    'I use this path, but maybe you need to change It.
    strSmiley = LoadArrayFile(App.Path & "\media\smileys\smileys.dat")
    Dim ni As Integer
    'Load each Smiley and remove invalid chars
    For ni = 0 To 78
        strSmiley(ni) = RemoveInvalidChars(strSmiley(ni))
    Next
    Exit Sub
NoSmileys:
    'An error, I'll close the app, but you can do some other things
    frmSplash.SetStatus "Smileys Not Found! Shuting Down..."
    MsgBox "The Smiley Definition File was not found! please reinstall the program." & vbCrLf & "The programm will End.", vbCritical
    Unload frmMain
End Sub

'Remove vbcrlf and other chars
Public Function RemoveInvalidChars(ByVal strSource As String) As String
    Dim ni As Integer
    For ni = 1 To Len(strSource)
        If Mid(strSource, ni, 1) < Chr(32) Then Mid(strSource, ni, 1) = " " ''Debug.Print Mid(strSource, ni, 1) '
    Next ni
    RemoveInvalidChars = Trim(strSource)
End Function

'In a text string (text only), replace the text equivalents for smileys
' with the html image tags
Public Function AddSmileys(sInput As String) As String
    Dim strFinalMessage As String
    Dim ni As Integer
    If bShowSmileys Then
        strFinalMessage = sInput
        'Copy String
        'Search Emoticons!
        For ni = 78 To 0 Step -1
            ' a temporal solution for uppercase and lowercase is do IT twice
            strFinalMessage = Replace(strFinalMessage, UCase(strSmiley(ni)), "<IMG src='" & App.Path & "\media\Smileys\" & ni + 1 & ".gif' align=top border=0>")
            'sugestions are wellcome
            strFinalMessage = Replace(strFinalMessage, LCase(strSmiley(ni)), "<IMG src='" & App.Path & "\media\Smileys\" & ni + 1 & ".gif' align=top border=0>")
        Next ni
        'replace the resulting string
        AddSmileys = strFinalMessage
    Else
        AddSmileys = sInput
    End If
End Function
