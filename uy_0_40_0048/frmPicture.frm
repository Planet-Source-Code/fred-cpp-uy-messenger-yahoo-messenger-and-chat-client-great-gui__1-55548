VERSION 5.00
Begin VB.Form frmPicture 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Preview"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   945
   Icon            =   "frmPicture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgPreview 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

