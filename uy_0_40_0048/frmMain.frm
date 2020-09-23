VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{2B323CCC-50E3-11D3-9466-00A0C9700498}#1.0#0"; "yacscom.dll"
Begin VB.Form frmMain 
   Caption         =   "UY! Messenger"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   195
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      _Version        =   327682
   End
   Begin InetCtlsObjects.Inet inetGrab 
      Left            =   1320
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsChat 
      Left            =   3120
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsYahoo 
      Left            =   3600
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6780
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12515
            Text            =   "Yahoo Status"
            TextSave        =   "Yahoo Status"
            Key             =   "status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Your Yahoo Status"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pChat 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   9015
      Begin ComctlLib.Toolbar tbTextFormat 
         Height          =   390
         Left            =   0
         TabIndex        =   4
         Top             =   3720
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "imlFormat"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   14
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Chat"
               Object.ToolTipText     =   "Change the room you are In"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Text"
               Object.ToolTipText     =   "Send the text as normal message"
               Object.Tag             =   ""
               ImageIndex      =   2
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Thought"
               Object.ToolTipText     =   "Send the message as . o 0 ( Thought )"
               Object.Tag             =   ""
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Emote"
               Object.ToolTipText     =   "Emote. Send text as an Emotion"
               Object.Tag             =   ""
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Set text to Bold style"
               Object.Tag             =   ""
               ImageIndex      =   5
               Style           =   1
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Set text to Italic Style"
               Object.Tag             =   ""
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Set text to Underline Style"
               Object.Tag             =   ""
               ImageIndex      =   7
               Style           =   1
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Color"
               Object.ToolTipText     =   "Set the Text Color"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Smileys"
               Object.ToolTipText     =   "Insert Smileys!"
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Audio"
               Object.ToolTipText     =   "add audio to this room"
               Object.Tag             =   "Audio"
               ImageIndex      =   10
               Style           =   1
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
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
            Left            =   3840
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
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   975
         End
      End
      Begin ComctlLib.Toolbar tbAudio 
         Height          =   390
         Left            =   0
         TabIndex        =   14
         Top             =   4080
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "imlisebItems"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   1
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Lock"
               Object.ToolTipText     =   "Press to lock the microphone"
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
         EndProperty
         Begin VB.CommandButton cmdTalk 
            Caption         =   "Talk"
            Height          =   315
            Left            =   720
            TabIndex        =   15
            Top             =   30
            Width           =   1215
         End
         Begin YACSCOMLibCtl.YAcs yacsChat 
            Left            =   6960
            OleObjectBlob   =   "frmMain.frx":1CFA
            Top             =   0
         End
      End
      Begin VB.TextBox txtMessage 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         MaxLength       =   800
         TabIndex        =   10
         Top             =   4440
         Width           =   6135
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6240
         TabIndex        =   9
         Top             =   4440
         Width           =   1095
      End
      Begin SHDocVwCtl.WebBrowser wbCHat 
         Height          =   3615
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   4815
         ExtentX         =   8493
         ExtentY         =   6376
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
         Location        =   "http:///"
      End
      Begin ComctlLib.TreeView tvCHaters 
         Height          =   3615
         Left            =   4800
         TabIndex        =   8
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6376
         _Version        =   327682
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pMain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   11
      Top             =   960
      Width           =   7695
      Begin uy.isExplorerBar isebMain 
         Height          =   4815
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   8493
         FontCharset     =   0
      End
      Begin ComctlLib.TreeView tvFriends 
         Height          =   4815
         Left            =   3240
         TabIndex        =   13
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8493
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   423
         LabelEdit       =   1
         Style           =   1
         ImageList       =   "imlStatus"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tsMain 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9551
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Friends"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Chat"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imlFormat 
      Left            =   2520
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2070
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":23C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2714
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":310A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":345C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":37AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3D00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlStatus 
      Left            =   1920
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4052
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4164
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4276
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4388
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":48DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":49EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4C10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlisebItems 
      Left            =   120
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5074
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":53C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":610E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6460
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":67B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":71A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlMain 
      Left            =   720
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   34
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":74FA
            Key             =   "Mesesage"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":886C
            Key             =   "Chat"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":9BDE
            Key             =   "Invite"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":AF50
            Key             =   "WebCam"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":C2C2
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D634
            Key             =   "File"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":E9A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":FD18
            Key             =   "Block"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1108A
            Key             =   "Mobile"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":11B3C
            Key             =   "Voice"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":12EAE
            Key             =   "No Voice"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnlogin 
      Caption         =   "login"
      Visible         =   0   'False
   End
   Begin VB.Menu mntest 
      Caption         =   "test"
      Visible         =   0   'False
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'YCHT Object
Public WithEvents YCht As Protocol_YCHT
Attribute YCht.VB_VarHelpID = -1

' Public Variables for yahoo
Public sYUser As String
Public sYPass As String
Public sYRoom As String
Public sYAlias As String

'Local Variables
Dim Crypt(1) As String
Dim lPrevChathWnd As Long
Dim bChatAttached As Boolean
Dim bScrollChatWindow As Boolean
Dim imsgMode As Integer
Dim lpCount As Long
Dim bClickAllowed As Boolean
Dim iPreviousITem As Integer
Dim bDisconectedByUser As Boolean

'Local Variables For Text Style in chat
Public iChatFontSize As Integer
Public sChatFontName As String
Public sChatFontColor As String
Public bChatFontBold As Boolean
Public bChatFontItalic As Boolean
Public bChatFontUnderline As Boolean

'More control variables
Public rKey As String, rm_Space As String


'*******************************************
'
' EVENT HANDLER

Private Sub cmdSend_Click()
    If Len(txtMessage.Text) > 0 Then
        Select Case imsgMode
            Case 0  'Normal Message
                YCht.ychtSendMessage AssambleMessage(txtMessage.Text)
            Case 1  'Thought
                YCht.ychtSendEmote AssambleThought(AddSmileys(txtMessage.Text))
            Case 2  'Emote
                YCht.ychtSendEmote " " & AddSmileys(txtMessage.Text)
        End Select
        txtMessage.Text() = vbNullString
    End If
End Sub

Private Sub comFontName_Click()
    sChatFontName = comFontName.Text
    txtMessage.FontName = comFontName.Text
End Sub

Private Sub comSize_Click()
    iChatFontSize = comSize.Text
    txtMessage.FontSize = iChatFontSize
End Sub

Private Sub Form_Load()
    'Ceck for multi Instance
    ChDir App.Path
    If App.PrevInstance Then
        If MsgBox("A previous instance was create, create a new Instance?", vbYesNo + vbQuestion) = vbNo Then End
    End If
    frmSplash.Show
    frmSplash.SetStatus "Initializing..."
    'Clear up the document
    bClickAllowed = True
    wbChat.Navigate "about:Blank"
    frmSplash.SetStatus "Initializing chat system..."
notReady:
    DoEvents
    If wbChat.Busy Then GoTo notReady
    frmSplash.SetStatus "Initializing protocol..."
    bClickAllowed = False
    'Initialize our class module
    frmSplash.SetStatus "Initializing YCHT Class module..."
    Set frmMain.YCht = New Protocol_YCHT
    frmMain.YCht.GetSock wsChat
    'Load Settings
    frmSplash.SetStatus "Loading settings..."
    'Setup Login Variables
    ReadSettings
    sYUser = ""
    sYPass = ""
    sYRoom = "Chat Central:1"
    'Options Variables
    bAceptIMs = 1
    bScrollChatWindow = True
    imsgMode = 0
    'Init isExplorerBar Structure
    frmSplash.SetStatus "Creating ExplorerBar Structure..."
    BuildBarStructure
    'Load Fonts
    frmSplash.SetStatus "Loading Fonts..."
    Dim ni As Integer
    For ni = 0 To Screen.FontCount
        comFontName.AddItem Screen.Fonts(ni)
    Next ni
    'Load Smileys
    frmSplash.SetStatus "Loading Smileys..."
    LoadSmileys
    'Load Font Sizes
    For ni = 6 To 32
        comSize.AddItem ni
    Next ni
    'Setup Font Variables
    iMyFontSize = 10
    sMyFontName = "Arial"
    sMyFontColor = "#000000"
    bMyFontBold = False
    bMyFontItalic = False
    bMyFontUnderline = False
    'Local Copy for chat
    'Local Variables For Text Style in chat
    iChatFontSize = iMyFontSize
    sChatFontName = sMyFontName
    sChatFontColor = sMyFontColor
    bChatFontBold = bMyFontBold
    bChatFontItalic = bMyFontItalic
    bChatFontUnderline = bMyFontUnderline
    
    comFontName.Text = sMyFontName
    comSize.Text = iMyFontSize
    txtMessage.FontName = sMyFontName
    txtMessage.FontSize = iMyFontSize
    isebMain.Font.name = "Verdana"
    frmSplash.SetStatus "Adding App data..."
    'wbCHat.Document.script.Document.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & "<p align=""center""><b><font face=""Verdana"" color=""#000080"" size=""2"">UY! Messenger</font></b><br><font size=""1""><B>http://www.geocities.com/uy_messenger/</b></font></p>"
    wbChat.Document.script.Document.write "<p align=""center""><b><font face=""Verdana"" color=""#000080"" size=""2"">UY! Messenger</b></font></p>"
    Set Me.tbMain.ImageList = Me.imlMain
    Dim tmp As Integer
    frmSplash.SetStatus "Creating toolbar..."
    For tmp = 1 To 5 'Me.imlMain.ListImages.Count
        tbMain.Buttons.Add , imlMain.ListImages(tmp).key, , , tmp
        'disable items that don't work [almost all them :( ]
        tbMain.Buttons(imlMain.ListImages(tmp).key).Enabled = False
    Next tmp
    'Enable working Items
    ' Show About
    tbMain.Buttons("Chat").Style = tbrCheck
    Show
    frmSplash.SetStatus "Finishing..."
    DoEvents
    frmSplash.Hide
    mnlogin_Click
End Sub

Private Sub Form_Resize()
    'Resize controls In the Form
    DoEvents
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    tsMain.Move 2, tbMain.Height + 2, ScaleWidth - 4, ScaleHeight - 4 - tbMain.Height - sbStatus.Height
    pChat.Move 7, tsMain.Top + 24, ScaleWidth - 14, tsMain.Height - 32
    pMain.Move 7, tsMain.Top + 24, ScaleWidth - 14, tsMain.Height - 32
    'Friends Controls
    isebMain.Move 0, 0, isebMain.Width, pMain.ScaleHeight
    tvFriends.Move isebMain.Width + 2, 1, pMain.ScaleWidth - isebMain.Width - 4, pMain.ScaleHeight - 2
    'Chat controls
    If tbAudio.Visible Then
        wbChat.Move 0, 0, ScaleWidth - 18 - tvCHaters.Width, pChat.ScaleHeight - tbTextFormat.Height - tbAudio.Height - txtMessage.Height - 6
        tvCHaters.Move wbChat.Width + 2, 0, tvCHaters.Width, wbChat.Height
        tbTextFormat.Move 0, wbChat.Height + 2, ScaleWidth - 14
        tbAudio.Move 0, wbChat.Height + tbTextFormat.Height + 2, ScaleWidth - 14
        txtMessage.Move 2, wbChat.Height + tbTextFormat.Height + tbAudio.Height + 4, ScaleWidth - 18 - cmdSend.Width
        cmdSend.Move ScaleWidth - cmdSend.Width - 15, txtMessage.Top
    Else
        wbChat.Move 0, 0, ScaleWidth - 18 - tvCHaters.Width, pChat.ScaleHeight - tbTextFormat.Height - txtMessage.Height - 6
        tvCHaters.Move wbChat.Width + 2, 0, tvCHaters.Width, wbChat.Height
        tbTextFormat.Move 0, wbChat.Height + 2, ScaleWidth - 14
        txtMessage.Move 2, wbChat.Height + tbTextFormat.Height + 4, ScaleWidth - 18 - cmdSend.Width
        cmdSend.Move ScaleWidth - cmdSend.Width - 15, txtMessage.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub isebMain_ItemClick(sGroup As String, sItemKey As String)
    Select Case sItemKey
    '***************************************
    ' Main UY! Messenger Items
    Case "Home"
        'Visit homepage
        isebMain.OpenLink "http://www.geocities.com/uy_messenger/"
    Case "Update"
        'Check For Updates
        frmPreferences.ShowSection "Updates"
    Case "Login"
        'Login/Change User
        mnlogin_Click
    Case "About"
        'Show About Box
        mnAbout_Click
    
    '***************************************
    ' More Options
    Case "Preferences"
        'Start Preferences Dialog
        frmPreferences.ShowSection "Privacity"
    Case "ChatZero"
        'Start Chat0
         MsgBox "this option is under development. " & vbCrLf & _
                "Chat0 (-ChatZero-) is a project of an Unbotable " & vbCrLf & _
                "chat client I'm working on. I'll update this soon", vbInformation
    '***************************************
    ' Selected User Comands
    Case "IM"
        'Send An Instant Message
        PMUser tvFriends.SelectedItem.key
    Case "Voice"
        'Start Voice chat
        'Not Implemented
    Case "Cam"
        'Send My WebCam
        'Not Implemented
    Case "Conference"
        'Invite to conference
        'guess what..
    Case "Picture"
        'View User Photo
        frmPicture.imgPreview.Picture = LoadPicture(y_profile.LastImage)
        frmPicture.Height = frmPicture.imgPreview.Height + 300
        frmPicture.Width = frmPicture.imgPreview.Width
        frmPicture.Show , Me
    Case Else
        Debug.Print "Unrecogniced command: " & sGroup & " - " & sItemKey
    End Select
End Sub

Private Sub mnAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnlogin_Click()
    'Get user, pass and room
    frmLogin.Show vbModal
    If frmLogin.comID = "" Then Exit Sub
    tvFriends.Nodes.Clear
    sTmpFriends = ""
    sTmpIgnored = ""
    sTmpAlias = ""
    wsYahoo.Close
    wsYahoo.Connect Host_Name, 5050
    Unload frmLogin
End Sub

Private Sub mntest_Click()
    'I use this for multimple test, just Ignore It. =D
'    Dim ni As Integer
'    'Debug.Print "*****-comparacion"
'    For ni = 1 To 7
'        'Debug.Print ni & "=> "; FontSizeToPT(ni)
'    Next ni
'    'Debug.Print "*****-comparacion"
    
    If bChatAttached Then
        lPrevChathWnd = AttachWindow(frmChat, 0, False)
    Else
        lPrevChathWnd = AttachWindow(frmChat, pChat.hwnd, True)
    End If
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "Chat"
            bDisconectedByUser = False
            'Presionado: Estas conectado.
            If Button.Value = tbrUnpressed Then
                'Desconectar!
                bDisconectedByUser = True
                YCht.ychtClose
            Else
                frmRooms.Show
            End If
    End Select
End Sub

Private Sub tbTextFormat_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "Chat"
            frmRooms.Show vbModal, Me
        Case "Text"
            imsgMode = 0
        Case "Thought"
            imsgMode = 1
        Case "Emote"
            imsgMode = 2
        Case "Smileys"
            frmSmileys.GetSmiley Me
        Case "Bold"
            bChatFontBold = Button.Value
            txtMessage.FontBold = bChatFontBold
        Case "Italic"
            bChatFontItalic = Button.Value
            txtMessage.FontItalic = bChatFontItalic
        Case "Underline"
            bChatFontUnderline = Button.Value
            txtMessage.FontUnderline = bChatFontUnderline
        Case "Color"
            sChatFontColor = frmColor.GetColor(Me, sChatFontColor)
            txtMessage.ForeColor = frmColor.GetLastLColor
        Case "Audio"
            'Add audio to the chat room
            AddAudio
    End Select
End Sub

Private Sub tsMain_Click()
    Select Case tsMain.SelectedItem.Caption
        Case "Friends"
            pMain.Visible = True
            pChat.Visible = False
        Case "Chat"
            pChat.Visible = True
            pMain.Visible = False
    End Select
End Sub

Private Sub tvCHaters_DblClick()
    PMUser tvCHaters.SelectedItem.Text
End Sub

Private Sub tvFriends_Click()
    If Not tvFriends.SelectedItem Is Nothing Then
        If iPreviousITem <> tvFriends.SelectedItem.Index Then
            BuildBarStructure
            iPreviousITem = tvFriends.SelectedItem.Index
        End If
    End If
End Sub

Private Sub tvFriends_DblClick()
    If tvFriends.SelectedItem Is Nothing Then Exit Sub
    With tvFriends.SelectedItem
        If .key = "_Main" Then
        'Main Node
            'What To Do?
        ElseIf Left$(.key, 5) = "group" Then
        'Group Node
            If .Expanded Then
                .Image = 3 'Me.imlStatus.ListImages.Item(2)
            Else
                .Image = 2 'Me.imlStatus.ListImages.Item(1)
            End If
        Else
        'User Node - Instant Message
            PMUser .key
        End If
    End With
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSend_Click
End Sub

Private Sub wbCHat_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim sAction As String
    If bClickAllowed Then
        Cancel = False
    Else
        Cancel = True
        'about:blank/PM:usertobepmed
        sAction = Mid(URL, 13, 2)
        Select Case sAction
            Case "PM"
                'Send a PM to user
                PMUser Mid(URL, 16, Len(URL) - 15)
        End Select
    End If
End Sub

Private Sub wbCHat_GotFocus()
    bScrollChatWindow = False
End Sub

Private Sub wbCHat_LostFocus()
    bScrollChatWindow = True
End Sub

Private Sub wsYahoo_Connect()
    SetStatus "Logging in."
    wsYahoo.SendData Get_Key(sYUser)
End Sub

Private Sub wsYahoo_DataArrival(ByVal bytesTotal As Long)
    If bytesTotal = 0 Then GoTo StatDisconected
    Dim data As String, dData() As String, sdata() As String, i As Integer
    wsYahoo.GetData data, vbString, bytesTotal
    dData = Split(data, "YMSG" & Chr(0))
    For i = 0 To UBound(dData)
        data = "YMSG" & Chr(0) & dData(i)
        Select Case LCase(Mid(data, 12, 1))
            Case "w"    'Login
                sdata = Split(data, "À€")
                'getencrstrings sYUser, sYPass, sdata(3), Crypt(0), Crypt(1), 1
                getencrstrings sYPass, sdata(3), Crypt(0), Crypt(1)
                wsYahoo.SendData Login(sYUser, Crypt(0), Crypt(1))
            Case "u"    'Login info
                SetStatus "Logged in."
                tbMain.Buttons("Chat").Enabled = True
                GetYahooFriendsList (data)
            Case "" 'Login Initial Friends Status
                GetFriendsStatus data, tvFriends
            Case ""    'Set User Status (Aviable)
                SetFriendStatus data, tvFriends
            Case ""    'Set User Status (Bussy + id)
                SetFriendStatus data, tvFriends
            Case ""    'Set User Status (Offline)
                SetFriendOffline data, tvFriends
            Case ""    'Recibed PM
                If bAceptIMs <> 0 Then MPRecibed (data)
            Case "k"    ' Typing
                If bAceptIMs <> 0 Then TPMsg (data) 'wip
        End Select
        If Mid(data, 13, 4) = "ÿÿÿÿ" Then
            Disconect
        End If
    Debug.Print Replace(data, Chr(0), "*")
    Next
    Exit Sub
StatDisconected:
    Disconect
End Sub

Private Sub wsYahoo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description
End Sub

'*****************************************************
'
' LOCAL ROUTINES

'Set Status Text
Sub SetStatus(sNewStatus As String)
    sbStatus.Panels(1).Text = sNewStatus
End Sub

Public Sub BuildLists()
    Dim sGroupsArray() As String, tmpFriends() As String
    Dim tmpStr1, tmpStr2 As String, tmpstr3, tmpGroupName As String
    Dim ni As Integer, nj As Integer
    tvFriends.Visible = False
    tvFriends.Nodes.Clear
    DoEvents
    sGroupsArray = Split(sTmpFriends, Chr(10), , vbTextCompare)
    nj = 0
    'Build Friends List
    On Error Resume Next
    tvFriends.Nodes.Add , , "_Main", "Friends", 1, 1
    For Each tmpStr1 In sGroupsArray
        ni = InStr(1, tmpStr1, ":", vbTextCompare)
        If ni = 0 Then Exit For
        tmpGroupName = Left$(tmpStr1, ni - 1)
        'Add group
        tvFriends.Nodes.Add , , "group" & nj, tmpGroupName, 2, 2
        'Add friends
        tmpStr2 = Right$(tmpStr1, Len(tmpStr1) - ni)
        tmpFriends = Split(tmpStr2, ",", , vbTextCompare)
        For Each tmpstr3 In tmpFriends
            tvFriends.Nodes.Add "group" & nj, tvwChild, tmpstr3, tmpstr3, 4, 4
        Next
        tvFriends.Nodes.Item("group" & nj).Expanded = True
        nj = nj + 1
    Next
    'Build Ignore List
    For Each tmpStr1 In sGroupsArray
        ni = InStr(1, tmpStr1, ":", vbTextCompare)
        If ni = 0 Then Exit For
        tmpGroupName = Left$(tmpStr1, ni - 1)
        tmpStr2 = Right$(tmpStr1, Len(tmpStr1) - ni)
        'Debug.Print "Grupo: " & tmpGroupName
        'Debug.Print "Friends:" & tmpStr2
    Next
    'Build Friends List
    '    For Each tmpStr1 In sGroupsArray
    '        ni = InStr(1, tmpStr1, ":", vbTextCompare)
    '        If ni = 0 Then Exit For
    '        tmpGroupName = Left$(tmpStr1, ni - 1)
    '        tmpStr2 = Right$(tmpStr1, Len(tmpStr1) - ni)
    '        ''Debug.Print "Grupo: " & tmpGroupName
    '        ''Debug.Print "Friends:" & tmpStr2
    '    Next
    sAlias = Split(sTmpAlias, ",", 7, vbTextCompare)
    'Debug.Print "Aliases:" & sTmpAlias
ErrOcurred:
    tvFriends.Visible = True
    BuildBarStructure
    'list alreadi build
End Sub

Private Sub BuildBarStructure()
    'Here I'll add the structure for the isExplorerBar
    ' Depending of the selected item in the treeview
    ' Diferent Actions will be displayed.
    isebMain.SetDetailsImage
    isebMain.SetImageList imlisebItems
    If tvFriends.SelectedItem Is Nothing Then
        GoTo CreateAbout
    Else
        If tvFriends.SelectedItem.Parent Is Nothing Then
            GoTo CreateAbout
        Else
            isebMain.DisableUpdates True
            isebMain.SetImageList imlisebItems
            isebMain.ClearStructure
            With tvFriends.SelectedItem
                'Add Task (Special) group.
                isebMain.AddSpecialGroup "Actions for " & .key, Me.Icon
                'IM
                isebMain.AddItem -1, "IM", "Send an Instant Message", 8
                'Voice
                isebMain.AddItem -1, "Voice", "Start Voice Chat", 9
                'Cam
                isebMain.AddItem -1, "Cam", "Invite to see my cam", 10
                'Conference
                isebMain.AddItem -1, "Conference", "Invite to Conference", 11
                'Clear Image
                isebMain.SetDetailsImage
                'Add Details Text
                isebMain.AddDetailsGroup "Details", .key & " profile", "Searching...."
                'Let the control Refresh
                isebMain.DisableUpdates False
                'Get the profile
                GetAsyncProfile .key
            End With
        End If
        Exit Sub
    End If
CreateAbout:
    'Create the "about" structure.
    'Prevent Redrawing
    isebMain.DisableUpdates True
    'Clear previous groups and Items
    isebMain.ClearStructure
    'Create Special Group
    isebMain.AddSpecialGroup App.Title, frmAbout.imgIcon.Picture
    '
    isebMain.AddItem -1, "Home", "Go to uy! messenger home page", 1
    '
    isebMain.AddItem -1, "Update", "Check for updates", 2
    '
    isebMain.AddItem -1, "About", "About uy! messenger", 3
    isebMain.ExpandGroup -1, False
    'Add Extra options group
    isebMain.AddGroup "UYActions", "UY! Actions"
    isebMain.AddItem "UYActions", "Login", "Login", 5
    'Add Extra Options Items
    isebMain.AddItem "UYActions", "Preferences", "Show Preferences dialog", 6
    isebMain.AddItem "UYActions", "ChatZero", "Launch ChatZero", 7
    'Details Version
    isebMain.AddDetailsGroup "Details", "UY! Messenger", "Version: " & App.Major & "." & Format(App.Minor, "00") & vbCrLf & " Build: " & App.Revision & vbCrLf & vbCrLf & "Developed By Fred.cpp"
    'Set Icon
    isebMain.SetDetailsImage frmAbout.imgIcon.Picture
    'Let the control Refresh
    isebMain.DisableUpdates False
    'Relax
End Sub

Public Sub Disconect()
    'Disconect all services.
    tvFriends.Nodes.Clear
    tvFriends.Visible = False
    'if reconnect    ...
    If bAutoReConnect Then
        'If sProtocol = "YMSG11" Then
        'Using ymsg11 protocol
            SetStatus "You have been Disconected! Logging In back..."
            sTmpFriends = ""
            sTmpIgnored = ""
            sTmpAlias = ""
            wsYahoo.Close
            wsYahoo.Connect Host_Name, 5050
        'ElseIf sProtocol = "YCHT" Then
        'Using YCHT protocol
        '    sTmpFriends = ""
        '    sTmpIgnored = ""
        '    sTmpAlias = ""
        '    wsYahoo.Close
        '    wsYahoo.Connect Host_Name, 5050
        'End If
    Else
        SetStatus "Disconected!"
    End If
End Sub

Public Sub ymsgJoinRoom(sAlias As String, strRoom As String)
    'ShowNotification "Function not yet supported!"
    'Login in ycht protocol
    YCht.Server_Chat = "jcs1.chat.dcn.yahoo.com"
    'jcs1.chat.dcn.yahoo.com
    'scsc.msg.yahoo.com*
    'cs8.chat.yahoo.com
    'cs72.dcn.sc5.yahoo.com*
    YCht.Server_Login = "login.yahoo.com"
    YCht.Login_Username = sAlias
    frmMain.YCht.Login_Password = sYPass
    'Connect to YCHT Server
    frmMain.YCht.ychtConnect
    
End Sub

Private Function IsFriend(strUser As String) As Boolean
    On Error GoTo NoFriend
    If tvFriends.Nodes(strUser) Then
        IsFriend = True
    Else
        IsFriend = False
    End If
NoFriend:
    IsFriend = False
End Function

'*************************************************************
' HTML CODE EFFECTS
'*************************************************************
Private Sub AddChatText(sHTMLText As String)
    lpCount = lpCount + 1
    wbChat.Document.script.Document.write _
            "<DIV ID = ""DIV" & lpCount & """><B><font face='Verdana' color='#880000' size='1'>" & Format(Now, "hh:mm:ss") & ":</font></B>" & sHTMLText & "</i></b></u></a></DIV>" & vbCrLf
    If bScrollChatWindow Then wbChat.Document.script.Document.getElementById("DIV" & lpCount).scrollIntoView (True)
End Sub

Function AssambleMessage(strText As String) As String
    'Format Message to Be understood by the YCHT Protocol
    '[1m[4m[2m[34m<font face="fontname" size="fontsizept">Message
    AssambleMessage = _
                        IIf(bChatFontBold, "[1m", "") & _
                        IIf(bChatFontItalic, "[4m", "") & _
                        IIf(bChatFontUnderline, "[2m", "") & _
                        "[" & sChatFontColor & "m" & _
                        "<font face=""" & sChatFontName & _
                        """ size=""" & iChatFontSize & _
                        """>" & txtMessage.Text
End Function

Function AssambleThought(strText As String)
    'Format Message to looks like Thought yahelite Style
    '. o O ( Message )
    AssambleThought = " . o O ( " & strText & " )"
End Function

Public Sub ShowNotification(strNotification As String)
    AddChatText "<font name=""Verdana"" size=""2"" color=""#0000FF"">" & strNotification & "</font>"
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
            strcolor = "000000"
        End If
        'tmpStr = Mid(strText, 4, InStr(1, strText, "m", vbTextCompare) - 4)
        strText = Replace(strText, "[" & tmpStr & "m", "", , , vbTextCompare)
    Else
        strcolor = GetYTextColor(tmpStr)
    End If
    'Something Like This:
    '<font face="Paramount" size="17">TEXT_HERE''''17pt
    'Find Size
    ni = InStr(1, strText, "size=", vbTextCompare)
    If ni <> 0 Then
        tmpStr = Mid(strText, ni + 6, 2)
        If Mid(tmpStr, 2, 1) = """" Then tmpStr = Left$(tmpStr, 1)
        If Mid(tmpStr, 2, 1) = "'" Then tmpStr = Left$(tmpStr, 1)
        iFontSize = CInt(tmpStr)
    End If
    'Find Font
    ni = InStr(1, strText, "face=", vbTextCompare)
    If ni <> 0 Then
        tmpStr = Mid(strText, ni + 6, InStr(ni + 7, strText, """", vbTextCompare) - ni - 6)
        strFontName = tmpStr
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

' Description: add audio to the open Chat room
Private Sub AddAudio()
    If tbAudio.Visible Then
        tbAudio.Visible = True
    Else
        tbAudio.Visible = True
    End If
    Form_Resize
End Sub

Public Function SetUpVoice() As String
'    rKey = Split(data, "À€130À€")(1)
'    rKey = Split(rKey, "À€")(0)
'    rm_Space = Split(data, "À€129À€")(1)
'    rm_Space = Split(rm_Space, "À€")(0)
End Function

'*************************************************************
' CHAT CODE
'*************************************************************

Private Sub YCht_Away(strMsg As String)
    ShowNotification strMsg
End Sub

Private Sub YCht_Connected(IsConnected As Boolean)
    'Event is fired when our connection state changes
    '----
    'isConnected=true   : when connected
    'isConnected=false  : when disconnected
    If IsConnected = True Then
        YCht.ychtJoinRoom sYRoom
        tbMain.Buttons("Chat").Value = tbrPressed
    Else
        tbMain.Buttons("Chat").Value = tbrUnpressed
        ShowNotification "Disconected from yahoo!"
        pChat.Enabled = False
        If bAutoReConnect And Not bDisconectedByUser Then
            'Reconnect!
            ShowNotification "Reconecting to yahoo!"
            ymsgJoinRoom sYAlias, sYRoom
        Else
        End If
    End If
End Sub

Private Sub YCht_Error(strError As String)
    AddChatText "<font face=""Arial"" color=""#ff0000""><B>YCHT Error: " & strError & "</B></font>"
End Sub

Private Sub YCht_FriendStatus(strFriend As String, fStatus As FriendStatus)
    'Event is fired when a friend's status changes to Online/Offline/Chat/Games
    '----
    'I've used an Enum to handle the Statuses and should be self explanatory
    'in the below usage
    Select Case fStatus
        Case FriendStatus.ChatJoined
            ShowNotification strFriend & " in in chat. "
        Case FriendStatus.ChatLeft
            ShowNotification strFriend & " left chat. "
        Case FriendStatus.GamesJoined
            ShowNotification strFriend & " is in games. "
        Case FriendStatus.GamesLeft
            ShowNotification strFriend & " left games. "
        Case FriendStatus.OnlineFalse
            ShowNotification strFriend & " is offline. "
        Case FriendStatus.OnlineTrue
            ShowNotification strFriend & " is online. "
    End Select
End Sub

Private Sub ycht_ReceivedEmail(emailCount As String)
    'Event is fired when you receive a new e-mail on the name you're currently using
    '----
    'emailCount     : Count of how many emails you have
    'MsgBox "We've received a new e-mail! (total of " & emailCount & " email(s)).", vbInformation
    ShowNotification "We've received a new e-mail! (total of " & emailCount & " email(s)). "
End Sub

Private Sub YCht_ReceivedEmote(strUser As String, strMsg As String)
    AddChatText "<font face=""Arial"" color=""#800080"" size=""2"">" & strUser & "</B>" & strMsg & "</font>"
End Sub

Private Sub YCht_ReceivedInvite(strRoom As String, strUser As String)
    If Not bAceptIMs Then 'bIgnoreInvites Then
        AddChatText "<font face=""Arial"" color=""#800080"" size=""2"">" & strUser & "Has invited you to " & strRoom & " - Invitation Ignored "
    Else
        AddChatText "<font face=""Arial"" color=""#800080"" size=""2"">" & strUser & "Has invited you to " & strRoom & " - Click <a href = ""/Join:" & strRoom & """ >HERE</a> to Join "
    End If
End Sub

Private Sub YCht_ReceivedMessage(strUser As String, strMsg As String)
    AddChatText "<a href=""/PM:" & strUser & """><B><font face=""arial"" color=""#000000"" size=""2"">" & strUser & ": </font></B></a>" & FormatHTMLInput(strMsg) & "" 'No add <Br>
End Sub

Private Sub YCht_ReceivedPrivateMessage(strUser As String, strMsg As String)
    If bAceptIMs = 0 Then
        Exit Sub
    ElseIf bAceptIMs = 1 Then
        If IsFriend(strUser) Then
            AddChatText "<a href=""/PM:" & strUser & """><B><font face=""arial"" color=""#000080"" size=""1"">Private Message from " & strUser & ": </font></B></a>" & FormatHTMLInput(strMsg)
        Else
            'if is not a friend?
        End If
    Else    'Recibe all PM
        AddChatText "<a href=""/PM:" & strUser & """><B><font face=""arial"" color=""#000080"" size=""1"">Private Message from " & strUser & ": </font></B></a>" & FormatHTMLInput(strMsg)
    End If
End Sub

Private Sub YCht_RoomJoined(strRoom As String, strRoomTopic As String)
    pChat.Enabled = True
    tvCHaters.Nodes.Clear
    Caption = "UY! Messenger: " & " - " & strRoom
    AddChatText "<B><font face=""Arial"" size=""2"">Joined: - " & strRoom & " - " & strRoomTopic & "</b>"
    SetUpVoice
End Sub

Private Sub YCht_UserEntered(strUser As String)
    Dim i As Integer
    For i = tvCHaters.Nodes.Count To 1 Step -1
        If StrComp(strUser, tvCHaters.Nodes.Item(i), vbTextCompare) = 0 Then Exit Sub
    Next i
    tvCHaters.Nodes.Add , , strUser, strUser
    AddChatText "<font face=""arial"" color=""#008800"" size=""2"">" & strUser & " has joined the room.</font>"
End Sub

Private Sub YCht_UserLeft(strUser As String)
    Dim i As Integer
    For i = tvCHaters.Nodes.Count To 1 Step -1
        'If the user exists in our list then remove them
        '(suggested that you should compare each list(i) with strUser in lowercase)
        If StrComp(strUser, tvCHaters.Nodes.Item(i), vbTextCompare) = 0 Then
            tvCHaters.Nodes.Remove i
            AddChatText "<font face=""Arial"" color=""#aa0000"" size=""2"">" & strUser & " has left the room.</font>"
        End If
    Next i
End Sub
