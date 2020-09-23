VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Color"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRGB 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Text            =   "Hex[ 00,00,00 ]"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.HScrollBar hsColor 
      Height          =   255
      Index           =   2
      LargeChange     =   32
      Left            =   840
      Max             =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
   Begin VB.HScrollBar hsColor 
      Height          =   255
      Index           =   1
      LargeChange     =   32
      Left            =   840
      Max             =   255
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.HScrollBar hsColor 
      Height          =   255
      Index           =   0
      LargeChange     =   32
      Left            =   840
      Max             =   255
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel  "
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK "
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Select RGB Color"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "Color:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape shpRGB 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   2  'Dash
      Height          =   285
      Left            =   3000
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BLUE"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "GEEN"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RED"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   345
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   480
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   480
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNewColor As String

Public Function GetColor(ByRef fOwner As Form, sColor As String) As String
    Dim sR As String, sG As String, sB As String
    sR = Mid(sColor, 2, 2)
    sG = Mid(sColor, 4, 2)
    sB = Mid(sColor, 6, 2)
    hsColor(0).Value = Val("&H" & sR)
    hsColor(1).Value = Val("&H" & sG)
    hsColor(2).Value = Val("&H" & sB)
    Show vbModal, fOwner
    GetColor = sNewColor
End Function

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Hide
    sNewColor = "#" & Right(Format(Hex(hsColor(0).Value), "00"), 2) & Right(Format(Hex(hsColor(1).Value), "00"), 2) & Right(Format(Hex(hsColor(2).Value), "00"), 2)
    Hide
End Sub

Private Sub hsColor_Change(Index As Integer)
    hsColor_Scroll Index
End Sub

Private Sub hsColor_Scroll(Index As Integer)
    'Update RGB Color
    Dim color(0 To 2) As Integer
    color(0) = 0: color(1) = 0: color(2) = 0
    color(Index) = hsColor(Index).Value
    shpColor(Index).BackColor = RGB(color(0), color(1), color(2))
    shpRGB.BackColor = RGB(hsColor(0).Value, hsColor(1).Value, hsColor(2).Value)
    txtRGB.Text = "Hex[ " & Format(Hex(hsColor(0).Value), "00") & "," & Format(Hex(hsColor(1).Value), "00") & "," & Format(Hex(hsColor(2).Value), "00") & "]"
End Sub

Public Function GetLastLColor() As Long
    GetLastLColor = RGB(hsColor(0).Value, hsColor(1).Value, hsColor(2).Value)
End Function
