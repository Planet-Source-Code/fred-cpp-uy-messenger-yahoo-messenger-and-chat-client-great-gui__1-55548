VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About UY! Messenger"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcredits 
      Caption         =   "Credits"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.Timer timHideCredits 
         Interval        =   8000
         Left            =   5400
         Top             =   1080
      End
      Begin SHDocVwCtl.WebBrowser wbCredits 
         Height          =   1815
         Left            =   -15
         TabIndex        =   7
         Top             =   -15
         Visible         =   0   'False
         Width           =   4965
         ExtentX         =   8758
         ExtentY         =   3201
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UY! Messenger"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   120
         Width           =   3375
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Left            =   0
         Picture         =   "frmAbout.frx":1CFA
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed By Fred.cpp"
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
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version ##.## Build ###"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "This software is free. For suggestions, comments and Bug Reports please see the preferences seccion."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'
' Project Name: uy! messenger
'
' Author:       Fred.cpp
'               fred_cpp@msn.com
'               fred_cpp@yahoo.com.mx
'
' Page:         http://www.geocities.com/uy_messenger
'
' Current
' Version:      0.40 Build 48
'
' Description:  Before the yahoo! messenger Version 6.0 yahoo was
'               needing some gui improvements, version 5.x was ugly,
'               so I decided to try to make a nice looking chat client
'               Also, I wanted to have a unbotable client, so I started.
'               later, yahoo anounced the y!messenger v.6 and I stoped
'               this project. butt... maybe with some help this could
'               be a really nice client. so, Let's try.
'
' Features:     You can open multiple Instances of the app
'               (like Multiyahoo)
'
'               You can even logg in multiple nicknames from the
'               same account! try It.
'
'               Logs Into chat rooms with ycht protocol, and
'               Is as far as I know, Unbotable (In rooms, in messenger
'               you can fall)
'
'               * more. see webpage.
'
' Requeriments: Requires a dll from markus. you can get the file
'               from uy! home Page or from www.romanware.com
'               Tested on win 98 and XP
'
' Known Bugs:   Lots. but more than bugs are features not yet
'               implemented. feedback is wanted (also help)
'               when you enter a room, afther log out, you will
'               keep as "online"
'
' More Bugs:    If you found a Bug, please e-mail me.
'
' 2004:08:13    First Source Code Release.

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version: " & App.Major & "." & Format(App.Minor, "00") & " Build: " & App.Revision
    DoEvents
    wbCredits.Navigate "http://mx.geocities.com/uy_messenger/about.htm"
End Sub

Private Sub timHideCredits_Timer()
    timHideCredits.Enabled = False
    wbCredits.Visible = False
End Sub

Private Sub cmdClose_Click()
    Hide
End Sub

Private Sub cmdcredits_Click()
    wbCredits.Visible = True
    timHideCredits.Enabled = True
End Sub


