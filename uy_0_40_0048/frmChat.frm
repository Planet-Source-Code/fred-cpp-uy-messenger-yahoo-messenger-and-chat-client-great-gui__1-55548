VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmChat 
   Caption         =   "UY! Chat"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   327682
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   615
      Left            =   6840
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   6615
   End
   Begin ComctlLib.TreeView tvChatUsers 
      Height          =   3735
      Left            =   5400
      TabIndex        =   1
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6588
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin SHDocVwCtl.WebBrowser wbChat 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   6588
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
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

