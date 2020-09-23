VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPreferences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UY! Messenger Preferences"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet inetUpdate 
      Left            =   3240
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin uy.isExplorerBar isebPreferences 
      Align           =   3  'Align Left
      Height          =   4620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8149
      FontCharset     =   0
   End
   Begin VB.PictureBox pUpdates 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   3120
      ScaleHeight     =   3495
      ScaleWidth      =   5295
      TabIndex        =   2
      Tag             =   "Tab"
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "CheckFor Updates"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblUpdatesStatus 
         Caption         =   "Press the Check Button to Search Updates of UY! Messenger"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox pPrivacity 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   3120
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   7
      Tag             =   "Tab"
      Top             =   600
      Width           =   5295
      Begin VB.ListBox listIgnored 
         Height          =   1230
         ItemData        =   "frmPreferences.frx":2982
         Left            =   2400
         List            =   "frmPreferences.frx":2989
         TabIndex        =   16
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Private Messages"
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5055
         Begin VB.PictureBox pIMs 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   4815
            TabIndex        =   10
            Top             =   480
            Width           =   4815
            Begin VB.OptionButton optIMs 
               Caption         =   "No! I don't want to recibe private Messages"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   480
               Width           =   4815
            End
            Begin VB.OptionButton optIMs 
               Caption         =   "Only from Buddyes and those who I've contacted"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   12
               Top             =   240
               Width           =   4815
            End
            Begin VB.OptionButton optIMs 
               Caption         =   "Yes! recibe Private mesages from every one"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   4815
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Do You Want to recibe Private Messages?"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Label Label2 
         Caption         =   "You won't recibe messages from these users"
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblIgnored 
         Caption         =   "Ignored Users"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         BorderWidth     =   3
         X1              =   368
         X2              =   8
         Y1              =   120
         Y2              =   120
      End
   End
   Begin SHDocVwCtl.WebBrowser webObj 
      Height          =   3495
      Left            =   3120
      TabIndex        =   5
      Tag             =   "Tab"
      Top             =   600
      Width           =   5295
      ExtentX         =   9340
      ExtentY         =   6165
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
   Begin VB.PictureBox pMore 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   3120
      ScaleHeight     =   3495
      ScaleWidth      =   5295
      TabIndex        =   19
      Tag             =   "Tab"
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txtRoom 
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         Text            =   "Chat Central:1"
         Top             =   960
         Width           =   2895
      End
      Begin VB.CheckBox chkAutoReconnect 
         Caption         =   "Auto Reconnect when Booted"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   4695
      End
      Begin VB.CheckBox chkAutoJoinRoom 
         Caption         =   "Auto join this Room at start up:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   4095
      End
      Begin VB.CheckBox chkSmileys 
         Caption         =   "Show Smileys"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         Caption         =   "Extras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         BorderWidth     =   3
         X1              =   0
         X2              =   5280
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label3 
         Caption         =   "Miscelaneous"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      X1              =   8400
      X2              =   3120
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H80000010&
      Caption         =   "Preferences"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpbAceptIMs As Integer

Private Sub chkAutoJoinRoom_Click()
    txtRoom.Enabled = CBool(chkAutoJoinRoom.Value)
    cmdApply.Enabled = True
End Sub

Private Sub chkAutoReconnect_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkSmileys_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    'Save changes
    bAceptIMs = tmpbAceptIMs
    bAutoJoinRoom = (IIf(chkAutoJoinRoom.Value, vbUnchecked, vbChecked))
    'chkAutoReconnect.Value = IIf(bAutoReConnect, vbUnchecked, vbChecked)
    bShowSmileys = IIf(chkSmileys.Value, vbChecked, vbUnchecked)
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    'Close
    Hide
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Save changes...
    cmdApply_Click
    'And close
    cmdCancel_Click
End Sub


Private Sub cmdUpdate_Click()
    If MsgBox("Sorry, this function is under development, Do you want to go to the UY! Messenger Home Page?", vbExclamation + vbYesNo, "UY! Messenger") = vbYes Then
        isebPreferences.OpenLink "http://www.geocities.com/uy_messenger/"
    End If
End Sub

Private Sub Form_Load()
    'Load Settings
    Dim sIgnored() As String
    Dim tmpItem
    optIMs(bAceptIMs).Value = True
    chkAutoJoinRoom.Value = IIf(bAutoJoinRoom, vbChecked, vbUnchecked)
    chkAutoReconnect.Value = IIf(bAutoReConnect, vbChecked, vbUnchecked)
    chkSmileys.Value = IIf(bShowSmileys, vbChecked, vbUnchecked)
    'Set Controls states
    chkAutoJoinRoom_Click
    cmdApply.Enabled = False
    'Load Ignored Users
    listIgnored.Clear
    sIgnored = Split(sTmpIgnored, ",")
    For Each tmpItem In sIgnored
        listIgnored.AddItem tmpItem
    Next
    'WIP: bUseCustomFonts
    
    'Preferences Structure:
    isebPreferences.AddSpecialGroup "Messenger", frmAbout.imgIcon
    isebPreferences.AddItem -1, "FeedBack", "Feedback", 1
    isebPreferences.AddItem -1, "History", "History", 2
    isebPreferences.AddItem -1, "Updates", "Updates", 3
    isebPreferences.AddGroup "Options", "Options"
    isebPreferences.AddItem "Options", "Privacity", "Privacity", 4
    isebPreferences.AddItem "Options", "More Options", "More Options", 5
End Sub

Private Sub isebPreferences_ItemClick(sGroup As String, sItemKey As String)
    Dim tmpCtl
    For Each tmpCtl In Me.Controls
        If tmpCtl.Tag = "Tab" Then tmpCtl.Visible = False
    Next
    Select Case sItemKey
        Case "FeedBack"
            webObj.Visible = True
            webObj.Navigate "http://mx.geocities.com/uy_messenger/mx_mail.htm"
        Case "History"
            webObj.Visible = True
            webObj.Navigate "http://mx.geocities.com/uy_messenger/mx_version_history.htm"
        Case "Updates"
            pUpdates.Visible = True
        Case "Privacity"
            pPrivacity.Visible = True
        Case "More Options"
            pMore.Visible = True
    End Select
    lblHeader.Caption = "  " & sItemKey
End Sub

Public Sub ShowSection(sKey As String)
    isebPreferences_ItemClick 0, sKey
    Show vbModal, frmMain
End Sub

Private Sub optIMs_Click(Index As Integer)
    tmpbAceptIMs = Index
    cmdApply.Enabled = True
End Sub

Private Sub webObj_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    'If Left(URL, 36) = "http://mx.geocities.com/uy_messenger" Then
    '    Cancel = False
    'Else
    '    Cancel = True
    'End If
End Sub

