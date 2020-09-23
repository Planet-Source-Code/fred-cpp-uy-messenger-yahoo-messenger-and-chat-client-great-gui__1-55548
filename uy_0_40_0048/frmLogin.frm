VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAutoJoinRoom 
      Caption         =   "Auto Join Room"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   5400
      Width           =   975
   End
   Begin VB.ComboBox comProtocol 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      TabIndex        =   12
      Top             =   3000
      Width           =   3255
   End
   Begin VB.ComboBox txtRoomLogin 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Text            =   "Chat Central:1"
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtRoomLogin_ 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CheckBox chkRememberID 
      Caption         =   "Remember my Password"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.ComboBox comID 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Protocol:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Room:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Login to Yahoo!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   120
      Picture         =   "frmLogin.frx":000C
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Yahoo ID"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMore

Private Sub chkRememberID_Click()
    'frmMain.Check2.Value = chkRememberID.Value
End Sub

Private Sub cmdCancel_Click()
    comID.Text = ""
    Hide
End Sub

Private Sub cmdLogin_Click()
    Dim strNames() As String
    Dim strRooms() As String
    Dim tmpStr As Variant
    Dim tmpStr2 As String
'    'Set the Chat and Login servers we'll be using
'    frmMain.YCht.Server_Chat = "jcs1.chat.dcn.yahoo.com"
'    'jcs1.chat.dcn.yahoo.com
'    'scsc.msg.yahoo.com*
'    'cs8.chat.yahoo.com
'    'cs72.dcn.sc5.yahoo.com*
'    frmMain.YCht.Server_Login = "login.yahoo.com"
'    'Change the below Username/Password to the one you wish to use
    frmMain.sYUser = comID.Text
    frmMain.sYPass = txtPassword.Text
    frmMain.sYRoom = txtRoomLogin.Text
    'Connect to YCHT Server
'    frmMain.YCht.ychtConnect
    'Add to Recent Names
    tmpStr2 = GetSetting("YCht", "Previous", "Names")
    strNames = Split(tmpStr2, ",")
    For Each tmpStr In strNames
        If tmpStr = comID.Text Then
            GoTo AlreadyExistsName
            'Dont save
        End If
    Next
    'Add entry
    SaveSetting "YCht", "Previous", "Names", tmpStr2 & "," & comID.Text
AlreadyExistsName:
    'Add to Recent Rooms
    tmpStr2 = GetSetting("YCht", "Previous", "Rooms")
    strRooms = Split(tmpStr, ",")
    For Each tmpStr In strRooms
        If tmpStr = comID.Text Then
            GoTo AlreadyExistsRoom
            'Dont save
        End If
    Next
    'Add entry
    SaveSetting "YCht", "Previous", "Rooms", tmpStr2 & "," & txtRoomLogin.Text
AlreadyExistsRoom:
    Hide
End Sub

Private Sub cmdMore_Click()
    Hide
    frmPreferences.ShowSection "More Options"
'    If bMore Then
'        Height = cmdMore.Top + cmdMore.Height + 520
'        cmdMore.Caption = "More options"
'    Else
'        Height = txtRoomLogin.Top + txtRoomLogin.Height + 520
'        cmdMore.Caption = "Fewer options"
'    End If
'    bMore = Not bMore
End Sub

Private Sub Form_Load()
    Dim strList() As String
    Dim ni As Integer
    Dim tmpStr As Variant
    bMore = True
    comID.Text = frmMain.sYUser
    txtPassword.Text = frmMain.sYPass
    txtRoomLogin.Text = frmMain.sYRoom
    'Populate combo with previous names
    strList = Split(GetSetting("YCht", "Previous", "Names"), ",")
    For Each tmpStr In strList
        comID.AddItem tmpStr
    Next
    'Populate combo with previous rooms
    strList = Split(GetSetting("YCht", "Previous", "Rooms"), ",")
    For Each tmpStr In strList
        txtRoomLogin.AddItem tmpStr
    Next
    'Populate Protocol
    comProtocol.AddItem "YMSG11"
    comProtocol.AddItem "YCHT"
    comProtocol.Text = "YCHT"
End Sub
