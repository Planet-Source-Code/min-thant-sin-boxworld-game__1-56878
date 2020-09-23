VERSION 5.00
Begin VB.Form frmSolved 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BoxWorld"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      TabIndex        =   0
      Top             =   1125
      Width           =   1590
   End
   Begin VB.Label lblMessage 
      Caption         =   "You've done a good job!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   825
      TabIndex        =   2
      Top             =   600
      Width           =   2940
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   225
      Picture         =   "frmSolved.frx":0000
      Top             =   225
      Width           =   450
   End
   Begin VB.Label lblMessage 
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Top             =   225
      Width           =   2955
   End
End
Attribute VB_Name = "frmSolved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
      On Error GoTo ErrorHandler
      
      Unload Me
      
      Call SaveGame
      frmLevels.File1.ListIndex = GameLevel
      LoadGame AddASlash(frmLevels.File1.Path) & frmLevels.File1.FileName
      
      Exit Sub
ErrorHandler:
      MsgBox Err.Description
      SaveGame
      Unload Me
End Sub
