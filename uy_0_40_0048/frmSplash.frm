VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   240
      Picture         =   "frmSplash.frx":000C
      Top             =   240
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
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version ##.## Build ###"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Height          =   615
      Left            =   -720
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    shpBorder.Move 0, 0, ScaleWidth, ScaleHeight
    shpBorder.Visible = True
    SetStatus "Initializing..."
    lblVersion.Caption = "Version: " & App.Major & "." & Format(App.Minor, "00") & " Build: " & App.Revision
End Sub

Public Sub SetStatus(strStatus As String)
    lblStatus.Caption = strStatus
    lblStatus.Refresh
    DoEvents
End Sub
