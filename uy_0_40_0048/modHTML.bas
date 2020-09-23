Attribute VB_Name = "modHTML"
Option Explicit

'Global Variables For Text Style
Public iMyFontSize As Integer
Public sMyFontName As String
Public sMyFontColor As String
Public bMyFontBold As Boolean
Public bMyFontItalic As Boolean
Public bMyFontUnderline As Boolean

'Extract Color from a string
Public Function GetYTextColor(strcolor As String) As String
    Select Case strcolor
        Case "31" 'Blue
            GetYTextColor = "#0000FF"
        Case "32" 'DarkCyan
            GetYTextColor = "#408080"
        Case "33" 'Gray
            GetYTextColor = "#808080"
        Case "34" 'green
            GetYTextColor = "#00FF00"
        Case "35" 'Fiusha
            GetYTextColor = "#FF0080"
        Case "36" 'Morado
            GetYTextColor = "#800080"
        Case "37" 'Orange
            GetYTextColor = "#FF8000"
        Case "38" 'red
            GetYTextColor = "#FF0000"
        Case "39" 'darkyellow
            GetYTextColor = "#888800"
        Case Else
            If Len(strcolor) = 0 Then
                GetYTextColor = "#000000"
            Else
                GetYTextColor = strcolor
            End If
    End Select
End Function

Public Function FontSizeToPT(iSize As Integer) As Integer
    'Traslate the size from 1-7 system to pixels size
    FontSizeToPT = (iSize) * (iSize - 2) * 0.6 + 10
End Function
 
Public Function FontSizeToIE(iSize As Integer) As Integer
    'Traslate the size from PixelSize to 1-7 system
    FontSizeToIE = (iSize / 3.7) - 1
End Function


'This works 8-| but need some optimization8-|
Public Function filter_HTML(strHTML As String) As String
    'The below just parses certain html tags from strings.. it's a bit rough but
    'it's not that important to optimize at the moment
'    Dim iPos1 As Integer, iPos2 As Integer, iCursor As Integer
'    ' Go into the string
'    For iCursor = 0 To Len(strHTML)
'        iPos1 = InStr(iCursor, strHTML, "", vbTextCompare)
'        If ipos Then    'Filter this code
'            iPos2 = InStr(iPos1, LCase(strHTML), "m")
'        End If
'    Next iCursor
    
    Dim Pos1 As Integer, Pos2 As Integer
reparse1:
    Pos1 = InStr(1, LCase(strHTML), "")
    If Pos1 > 0 Then
        Pos2 = InStr(Pos1, LCase(strHTML), "m")
        If Pos2 > 0 Then
            If Pos1 = 1 Then
                strHTML = Mid(strHTML, Pos2 + 1)
            Else
                strHTML = Mid(strHTML, 1, Pos1 - 1) & Mid(strHTML, Pos2 + 1)
            End If
        Else
            filter_HTML = strHTML
            GoTo reparse2
        End If
    Else
        filter_HTML = strHTML
        GoTo reparse2
    End If
    GoTo reparse1
reparse2:
    Pos1 = InStr(1, LCase(strHTML), "<font")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strHTML), "</")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strHTML), "</")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strHTML), "<b")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strHTML), "<alt")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strHTML), "<fade")
    If Pos1 > 0 Then
        Pos2 = InStr(Pos1, LCase(strHTML), ">")
        If Pos2 > 0 Then
            If Pos1 = 1 Then
                strHTML = Mid(strHTML, Pos2 + 1)
            Else
                strHTML = Mid(strHTML, 1, Pos1 - 1) & Mid(strHTML, Pos2 + 1)
            End If
        Else
           filter_HTML = strHTML
            GoTo ReadyToAddSmileys
        End If
    Else
        filter_HTML = strHTML
        GoTo ReadyToAddSmileys
    End If
    GoTo reparse2
ReadyToAddSmileys:
    filter_HTML = AddSmileys(filter_HTML)
End Function

