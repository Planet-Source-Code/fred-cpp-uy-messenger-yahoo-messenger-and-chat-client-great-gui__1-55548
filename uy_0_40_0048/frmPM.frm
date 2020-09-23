VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmPM 
   Caption         =   "Private Message"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6135
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tbTextFormat 
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlFormat"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Set text to Bold style"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Set text to Italic Style"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Set text to Underline Style"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Color"
            Object.ToolTipText     =   "Set the Text Color"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Smileys"
            Object.ToolTipText     =   "Insert Smileys!"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2000
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      Begin VB.ComboBox comFontName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   2415
      End
      Begin VB.ComboBox comSize 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
   End
   Begin ComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6135
      TabIndex        =   8
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox comAlias 
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "from:"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "To:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
   End
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   529
      SimpleText      =   "Uy! Messenger"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2725
            Text            =   "UY! Messenger"
            TextSave        =   "UY! Messenger"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   735
      Left            =   4920
      TabIndex        =   3
      Top             =   3195
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3195
      Width           =   4815
   End
   Begin SHDocVwCtl.WebBrowser wbCHat 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   3836
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin ComctlLib.ImageList imlFormat 
      Left            =   0
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":0C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":0F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":12D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":1624
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPM.frx":201A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'Global Variables For Text Style
Dim iFormFontSize As Integer
Dim sFormFontName As String
Dim sFormFontColor As String
Dim bFormFontBold As Boolean
Dim bFormFontItalic As Boolean
Dim bFormFontUnderline As Boolean

Dim bNavigate As Boolean
Private lpCount As Long

Private Sub cmdSend_Click()
    Dim sFinalMessage As String
    comAlias.Enabled = False
    txtTo.Enabled = False
    Tag = txtTo.Text
    sFinalMessage = AssambleMessage(txtMessage.Text)
    SendIM comAlias.Text, Tag, sFinalMessage
    AddChatText "<B><font face=""Arial"" color=""#000000"" size =""2"">" & frmMain.sYUser & ": </font></B>" & FormatHTMLOutput(txtMessage.Text)
    txtMessage.Text = ""
End Sub

Private Sub comFontName_Click()
    sFormFontName = comFontName.Text
    txtMessage.FontName = comFontName.Text
End Sub

Private Sub comSize_Click()
        iFormFontSize = comSize.Text
        txtMessage.FontSize = iFormFontSize
End Sub

Private Sub Form_Load()
    Dim tmp
    'Initialize HTML Object
    bNavigate = True
    wbCHat.Navigate "about:Blank"
notReady:
    DoEvents
    If wbCHat.Busy Then GoTo notReady
    iFormFontSize = iMyFontSize
    sFormFontName = sMyFontName
    sFormFontColor = sMyFontColor
    bFormFontBold = bMyFontBold
    bFormFontItalic = bMyFontItalic
    bFormFontUnderline = bMyFontUnderline
    'Load Fonts
    Dim ni As Integer
    For ni = 0 To Screen.FontCount
        comFontName.AddItem Screen.Fonts(ni)
    Next ni
    For ni = 6 To 32
        comSize.AddItem ni
    Next ni
    'Color
    sMyFontColor = "#000000"
    comFontName.Text = sFormFontName
    comSize.ListIndex = 4
    txtMessage.FontName = sFormFontName
    txtMessage.FontSize = iMyFontSize
    'Add alias for selection
    For Each tmp In sAlias
        comAlias.AddItem tmp
    Next
    comAlias.Text = IIf((frmMain.sYAlias = ""), frmMain.sYUser, frmMain.sYAlias)
    'wbCHat.Document.script.Document.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-16'>" & vbCrLf & "<p align=""center""><b><font face=""Verdana"" color=""#000080"" size=""2"">UY! Messenger</font></b></p>"
    wbCHat.Document.script.Document.write "<p align=""center""><b><font face=""Verdana"" color=""#000080"" size=""2"">UY! Messenger</font></b></p>"
    bNavigate = False
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    'Resize controls
    wbCHat.Move 2, tbMain.Height + 2, ScaleWidth - 3, ScaleHeight - tbMain.Height - tbTextFormat.Height - txtMessage.Height - sbStatus.Height - 12
    tbTextFormat.Move 2, wbCHat.Top + wbCHat.Height + 2, ScaleWidth - 3 ',  ScaleHeight - tbMain.Height - tbTextFormat.Height - txtMessage.Height - sbStatus.Height - 12
    txtMessage.Move 2, tbTextFormat.Top + tbTextFormat.Height + 2, ScaleWidth - cmdSend.Width - 6
    cmdSend.Move txtMessage.Width + 4, tbTextFormat.Top + tbTextFormat.Height + 2
End Sub

Private Sub AddChatText(sHTMLText As String)
    lpCount = lpCount + 1
    wbCHat.Document.script.Document.write _
            "<DIV ID = ""DIV" & lpCount & """><B><font face='Verdana' color='#880000' size='1'>" & Format(Now, "hh:mm:ss") & ":</font></B>" & sHTMLText & "</i></b></u></a></DIV>" & vbCrLf
    wbCHat.Document.script.Document.getElementById("DIV" & lpCount).scrollIntoView (True)
    Debug.Print Me.name
    Debug.Print Me.Caption
End Sub

Private Sub tbTextFormat_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "Smileys"
            frmSmileys.GetSmiley Me
        Case "Bold"
            bFormFontBold = Button.Value
            txtMessage.FontBold = bFormFontBold
        Case "Italic"
            bFormFontItalic = Button.Value
            txtMessage.FontItalic = bFormFontItalic
        Case "Underline"
            bFormFontUnderline = Button.Value
            txtMessage.FontUnderline = bFormFontUnderline
        Case "Color"
            sFormFontColor = frmColor.GetColor(Me, sFormFontColor)
            txtMessage.ForeColor = frmColor.GetLastLColor
    End Select
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = 13 Then
            txtMessage.Text = txtMessage.Text & vbCrLf
        ElseIf KeyCode = 71 Then
            SendDing
        End If
    Else
        If KeyCode = 13 Then
            cmdSend_Click
        End If
    End If
    
End Sub

Public Sub AddMessage(sMessage As String)
    If sMessage = "<ding>" Then
        AddDing
    Else
        AddChatText "<B><font face=""Arial"" color=""#0000FF"" size =""2"">" & Tag & ": </font></B>" & FormatHTMLInput(sMessage) ' & "<Br>"  '& FormatHTMLInput(strMsg) & "<Br>"
    End If
    sbStatus.Panels(1).Text = "Last message recibed at " & Now
End Sub

Public Sub AddStatusNotification(iType As Integer, Optional sNewStatus As String)
    Dim sStatusColor As String, sBackColor As String
    Select Case iType
        Case 4  'Offline
            sStatusColor = "#FF0000"
            sBackColor = "#FFF3F0"
        Case 5  'Online
            sStatusColor = "#008000"
            sBackColor = "#F2FFF2"
        Case 6  'Bussy
            sStatusColor = "#808080"
            sBackColor = "#DBEBFD"
        Case 7  'Away
            sStatusColor = "#808080"
            sBackColor = "#DBEBFD"
        Case 8  'Custom
            sStatusColor = ""
            sBackColor = "#F2FFF2"
    End Select
    lpCount = lpCount + 1
    wbCHat.Document.script.Document.write _
            "<DIV ID = ""DIV" & lpCount & """>" & _
            "<p style=""font-family: Verdana; font-size: 10pt; font-weight: bold; text-align: center;" & _
                "border: 1px solid #808080; padding-left: 4; padding-right: 4;" & _
                "padding-top: 1; padding-bottom: 1; background-color:" & sBackColor & """>" & _
                "<font color=""" & sStatusColor & """>" & _
                Format(Now, "hh:mm:ss") & ": " & Tag & " status is: " & sNewStatus & _
                "</font>" & _
            "<p></DIV>" & vbCrLf
    wbCHat.Document.script.Document.getElementById("DIV" & lpCount).scrollIntoView (True)
End Sub

Public Sub SendDing()
    Dim sFinalMessage As String
    comAlias.Enabled = False
    txtTo.Enabled = False
    Tag = txtTo.Text
    SendIM comAlias.Text, Tag, "<ding>"
    AddDing
End Sub

Public Sub AddDing()
    lpCount = lpCount + 1
    wbCHat.Document.script.Document.write _
            "<DIV ID = ""DIV" & lpCount & """><B><font face='Verdana' color='#FF0000' size='2'>BUZZ!!!</i></b></u></a></DIV>" & vbCrLf
    wbCHat.Document.script.Document.getElementById("DIV" & lpCount).scrollIntoView (True)
End Sub

Public Sub Flash()
    FlashWindow Me.hwnd, 0
End Sub

' This Function Analizes the input string and gets:
' Color, font, text atributes ( B, I, U ) and Size
Function FormatHTMLInput(strText As String) As String
    'Format Input Message to show HTML Text
    Dim bBold As Boolean
    Dim bUnderline As Boolean
    Dim bItalic As Boolean
    Dim tmpStr As String
    Dim strcolor As String
    Dim iFontSize As Integer
    Dim strFontName As String
    Dim sSimpleText As String
    Dim ni As Integer
    '/pm fred_cpp [1m[4m[2m[30mB_I_U_BLACK
    Const styleBold = "[1m"        '<!--Bold      -->
    Const styleUnderline = "[4m"   '<!--Underline -->
    Const styleItalic = "[2m"      '<!--Italic    -->
    '[3Xm <!--Color     -->
    If InStr(1, strText, styleBold, vbTextCompare) <> 0 Then bBold = True
    If InStr(1, strText, styleUnderline, vbTextCompare) <> 0 Then bUnderline = True
    If InStr(1, strText, styleItalic, vbTextCompare) <> 0 Then bItalic = True
    strText = Replace(strText, "[1m", "", , , vbTextCompare)
    strText = Replace(strText, "[4m", "", , , vbTextCompare)
    strText = Replace(strText, "[2m", "", , , vbTextCompare)
    'Find Color
    If Left$(strText, 2) = "[" Then
        'We Need to process the color.
        'GetText Value
        tmpStr = Mid(strText, 3, InStr(1, strText, "m", vbTextCompare) - 3)
        If Left$(tmpStr, 1) = "3" Then
            'Base Color (enumerated)
            strcolor = GetYTextColor(tmpStr)
        ElseIf Left$(tmpStr, 1) = "#" Then
            'RGBColor'[#0000a0m
            strcolor = tmpStr
        Else
            strText = filter_HTML(strText)
            strcolor = "#000000"
        End If
        'tmpStr = Mid(strText, 4, InStr(1, strText, "m", vbTextCompare) - 4)
        strText = Replace(strText, "[" & tmpStr & "m", "", , , vbTextCompare)
    Else
        strcolor = GetYTextColor(tmpStr)
    End If
    'Something Like This:
    'Find Size
    ni = InStr(1, strText, "size=", vbTextCompare)
    If ni <> 0 Then
        tmpStr = Mid(strText, ni + 6, 2)
        If Mid(tmpStr, 2, 1) = """" Then tmpStr = Left$(tmpStr, 1)
        If Mid(tmpStr, 2, 1) = "'" Then tmpStr = Left$(tmpStr, 1)
        iFontSize = CInt(tmpStr)
    Else
        iFontSize = 10
    End If
    'Find Font
    ni = InStr(1, strText, "face=", vbTextCompare)
    If ni <> 0 Then
        tmpStr = Mid(strText, ni + 6, InStr(ni + 7, strText, """", vbTextCompare) - ni - 6)
        strFontName = tmpStr
    Else
        strFontName = "Arial"
    End If
    ''Debug.Print "<font face=""" & strFontName & """ size=""" & iFontSize & "pt"">"
    FormatHTMLInput = IIf(bBold, "<b>", "") & _
                    IIf(bItalic, "<i>", "") & _
                    IIf(bUnderline, "<u>", "") & _
                    "<Font face='" & strFontName & "' style=""font-size: " & iFontSize & "pt""  color='" & strcolor & "'>" & _
                    filter_HTML(strText) & _
                    "</font>" & _
                    IIf(bBold, "</b>", "") & _
                    IIf(bItalic, "</i>", "") & _
                    IIf(bUnderline, "</u>", "")

End Function

' This Function Formats the html To Be shown in the own window
Function FormatHTMLOutput(strText As String) As String
    'Format Input Message to show HTML Text
    FormatHTMLOutput = IIf(bFormFontBold, "<b>", "") & _
                    IIf(bFormFontItalic, "<i>", "") & _
                    IIf(bFormFontUnderline, "<u>", "") & _
                    "<Font face='" & sFormFontName & "' size='" & FontSizeToIE(iFormFontSize) & "' color='" & sFormFontColor & "'>" & _
                    AddSmileys(strText) & _
                    "</font>" & _
                    IIf(bFormFontBold, "</b>", "") & _
                    IIf(bFormFontItalic, "</i>", "") & _
                    IIf(bFormFontUnderline, "</u>", "")

End Function

Function AssambleMessage(strText As String) As String
    'Format Message to Be understood by the YCHT Protocol
    '[1m[4m[2m[34m<font face="fontname" size="fontsizept">Message
    AssambleMessage = _
                        IIf(bFormFontBold, "[1m", "") & _
                        IIf(bFormFontItalic, "[4m", "") & _
                        IIf(bFormFontUnderline, "[2m", "") & _
                        "[" & sFormFontColor & "m" & _
                        "<font face=""" & sFormFontName & _
                        """ size=""" & iFormFontSize & _
                        """>" & txtMessage.Text
End Function


Private Sub wbCHat_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If Not bNavigate Then Cancel = True
End Sub

Public Sub Typing()
    sbStatus.Panels(1).Text = Tag & " is typing a message..."
End Sub
