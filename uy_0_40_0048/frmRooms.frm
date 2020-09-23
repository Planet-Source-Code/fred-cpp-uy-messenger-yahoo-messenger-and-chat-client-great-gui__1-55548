VERSION 5.00
Begin VB.Form frmRooms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Room"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRooms.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox comAlias 
      Height          =   360
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   3255
   End
   Begin VB.ComboBox comProtocol 
      Enabled         =   0   'False
      Height          =   360
      Left            =   3480
      TabIndex        =   5
      Text            =   "YCHT"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtRoom 
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ListBox listRooms 
      Height          =   1260
      ItemData        =   "frmRooms.frx":000C
      Left            =   120
      List            =   "frmRooms.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Join as:"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Protocol"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Recent Rooms"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Hide
    Unload Me
    If frmMain.YCht.IsConnected Then
        frmMain.tbMain.Buttons("Chat").Value = tbrPressed
    Else
        frmMain.tbMain.Buttons("Chat").Value = tbrUnpressed
    End If
End Sub

Private Sub cmdJoin_Click()
    Dim tmpStr As Variant
    Dim tmpStr2 As String
    Dim strRooms() As String
    frmMain.sYAlias = comAlias.Text
    frmMain.sYRoom = txtRoom.Text
    'If sProtocol = "YCHT" Then
    '    frmMain.YCht.ychtJoinRoom txtRoom.Text
    'Else
        frmMain.ymsgJoinRoom frmMain.sYAlias, frmMain.sYRoom
    'End If
    'Add to Recent Rooms
    tmpStr2 = GetSetting("YCht", "Previous", "Rooms")
    strRooms = Split(tmpStr, ",")
    For Each tmpStr In strRooms
        If tmpStr = txtRoom.Text Then
            GoTo AlreadyExistsRoom
            'Dont save
        End If
    Next
    'Add entry
    SaveSetting "YCht", "Previous", "Rooms", tmpStr2 & "," & txtRoom.Text
AlreadyExistsRoom:
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    Dim strList() As String
    Dim ni As Integer
    Dim tmpStr As Variant
    txtRoom.Text = frmMain.sYRoom
    'Populate List with previous rooms
    strList = Split(GetSetting("YCht", "Previous", "Rooms"), ",")
    For Each tmpStr In strList
        listRooms.AddItem tmpStr
    Next
    On Error GoTo NoAlias
    For Each tmpStr In sAlias
        comAlias.AddItem tmpStr
    Next
NoAlias:
    comAlias.Text = frmMain.sYUser
End Sub

Private Sub listRooms_Click()
    txtRoom.Text = listRooms.Text
End Sub

Private Sub listRooms_DblClick()
    cmdJoin_Click
End Sub
