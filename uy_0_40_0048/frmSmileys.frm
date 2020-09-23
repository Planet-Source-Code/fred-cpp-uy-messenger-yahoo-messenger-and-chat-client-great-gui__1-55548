VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmSmileys 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Insert Smiley"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSmileys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar tbSmileys 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imlSmileys 
      Left            =   1200
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   48
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":000C
            Key             =   ""
            Object.Tag             =   ":)"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":05C6
            Key             =   ""
            Object.Tag             =   ":("
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":0B80
            Key             =   ""
            Object.Tag             =   ";)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":113A
            Key             =   ""
            Object.Tag             =   ":D"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":16F4
            Key             =   ""
            Object.Tag             =   ";;)"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":1CAE
            Key             =   ""
            Object.Tag             =   ">:D<"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":2268
            Key             =   ""
            Object.Tag             =   ":-/"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":2822
            Key             =   ""
            Object.Tag             =   ":X"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":2DDC
            Key             =   ""
            Object.Tag             =   ":"">"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":3396
            Key             =   ""
            Object.Tag             =   ":P"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":3950
            Key             =   ""
            Object.Tag             =   ":-*"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":3F0A
            Key             =   ""
            Object.Tag             =   "=(("
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":44C4
            Key             =   ""
            Object.Tag             =   ":o"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":4A7E
            Key             =   ""
            Object.Tag             =   "x-("
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":5038
            Key             =   ""
            Object.Tag             =   ":>"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":55F2
            Key             =   ""
            Object.Tag             =   "B-)"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":5BAC
            Key             =   ""
            Object.Tag             =   ":-s"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":6166
            Key             =   ""
            Object.Tag             =   "#:-S"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":6720
            Key             =   ""
            Object.Tag             =   ">:)"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":6CDA
            Key             =   ""
            Object.Tag             =   ":(("
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":7294
            Key             =   ""
            Object.Tag             =   ":))"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":784E
            Key             =   ""
            Object.Tag             =   ":|"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":7E08
            Key             =   ""
            Object.Tag             =   "/:)"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":83C2
            Key             =   ""
            Object.Tag             =   "=))"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":897C
            Key             =   ""
            Object.Tag             =   "O:-)"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":8F36
            Key             =   ""
            Object.Tag             =   ":-B"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":94F0
            Key             =   ""
            Object.Tag             =   "=;"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":9AAA
            Key             =   ""
            Object.Tag             =   "I-)"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":A064
            Key             =   ""
            Object.Tag             =   "8-|"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":A61E
            Key             =   ""
            Object.Tag             =   "L-)"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":ABD8
            Key             =   ""
            Object.Tag             =   ":-&"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":B192
            Key             =   ""
            Object.Tag             =   ":-$"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":B74C
            Key             =   ""
            Object.Tag             =   "[-("
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":BD06
            Key             =   ""
            Object.Tag             =   ":o)"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":C2C0
            Key             =   ""
            Object.Tag             =   "8-}"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":C87A
            Key             =   ""
            Object.Tag             =   "<:-p"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":CE34
            Key             =   ""
            Object.Tag             =   "(:|"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":D3EE
            Key             =   ""
            Object.Tag             =   "=P~"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":D9A8
            Key             =   ""
            Object.Tag             =   ":-?"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":DF62
            Key             =   ""
            Object.Tag             =   "#-o"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":E51C
            Key             =   ""
            Object.Tag             =   "=D>"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":EAD6
            Key             =   ""
            Object.Tag             =   ":-ss"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":F090
            Key             =   ""
            Object.Tag             =   "@-)"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":F64A
            Key             =   ""
            Object.Tag             =   ":^o"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":FC04
            Key             =   ""
            Object.Tag             =   ":-w"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":101BE
            Key             =   ""
            Object.Tag             =   ":-<"
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":10778
            Key             =   ""
            Object.Tag             =   ">:p"
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSmileys.frx":10D32
            Key             =   ""
            Object.Tag             =   "<):)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSmileys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ParentForm As Form 'frmPM
Private Sub Form_Deactivate()
    Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Hide
End Sub

Private Sub Form_Load()
    Dim n As Integer
    Set tbSmileys.ImageList = imlSmileys
    For n = 1 To imlSmileys.ListImages.Count
        tbSmileys.Buttons.Add n, imlSmileys.ListImages(n).Tag, , , n
    Next n
End Sub

Private Sub Form_LostFocus()
    'Hide
End Sub

Private Sub Form_Resize()
    DoEvents
    Height = Screen.TwipsPerPixelY * (tbSmileys.Height + 25)
End Sub

Private Sub tbSmileys_ButtonClick(ByVal Button As ComctlLib.Button)
    Hide
    ParentForm.txtMessage.SelText = Button.key
End Sub

Public Function GetSmiley(ByRef objForm As Form) As String
    Set ParentForm = objForm
    Show , objForm
End Function
