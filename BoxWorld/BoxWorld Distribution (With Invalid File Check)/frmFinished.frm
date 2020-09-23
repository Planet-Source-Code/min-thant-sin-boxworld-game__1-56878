VERSION 5.00
Begin VB.Form frmFinished 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BoxWorld"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1425
      TabIndex        =   4
      Top             =   2025
      Width           =   1515
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3375
      TabIndex        =   3
      Top             =   2025
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "You have finished the last puzzle. Would you like to start over?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   675
      TabIndex        =   2
      Top             =   1050
      Width           =   5265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   840
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Width           =   5955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   75
      Width           =   5955
   End
End
Attribute VB_Name = "frmFinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
      Call SaveGame
      Unload Me
End Sub

Private Sub cmdYes_Click()
      On Error GoTo ErrorHandler
            
      Unload Me
      
      GameLevel = 0
      Call SaveGame
      
      frmLevels.File1.ListIndex = GameLevel
      LoadGame AddASlash(frmLevels.File1.Path) & frmLevels.File1.FileName
      
      Exit Sub
ErrorHandler:
      MsgBox "Couldn't load game level file"
End Sub
