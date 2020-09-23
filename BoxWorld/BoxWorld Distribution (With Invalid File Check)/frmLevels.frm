VERSION 5.00
Begin VB.Form frmLevels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Levels"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   4290
   End
   Begin VB.PictureBox picTmpBrowser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   5250
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.PictureBox picObjectsBrowser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   5250
      Picture         =   "frmLevels.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.PictureBox picBoardBrowser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6300
      Left            =   4425
      ScaleHeight     =   418
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   7200
   End
   Begin VB.DirListBox Dir1 
      Height          =   4590
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   450
      Width           =   4290
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   75
      Pattern         =   "*.bxw"
      TabIndex        =   0
      Top             =   5100
      Width           =   4290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8325
      TabIndex        =   2
      Top             =   6450
      Width           =   3315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Go to selected world"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4425
      TabIndex        =   1
      Top             =   6450
      Width           =   3690
   End
End
Attribute VB_Name = "frmLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private boolInvalidFile As Boolean

Private Sub cmdCancel_Click()
      Me.Hide
End Sub

Private Sub cmdOK_Click()
      If boolInvalidFile Then
            MsgBox "Choose another level file.", vbInformation
            Exit Sub
      End If
      
      Dim FileName As String
      
      FileName = AddASlash(File1.Path) & File1.FileName
      
      If Trim(File1.FileName) = "" Then Exit Sub
      
      Me.Hide
      
      LoadGame FileName
      GameLevel = File1.ListIndex
      Call SaveGame
End Sub

Private Sub Drive1_Change()
      On Error GoTo ErrorHandler
      Dir1.Path = Drive1.Drive
      Exit Sub
ErrorHandler:
      On Error Resume Next
      Drive1.Drive = Dir1.Path
End Sub

Private Sub File1_Click()
      DisplayLevel AddASlash(File1.Path) & File1.FileName
End Sub

Private Sub File1_DblClick()
      cmdOK_Click
End Sub

Private Sub Dir1_Change()
      On Error GoTo ErrorHandler
      File1.Path = Dir1.Path
      
      If File1.ListCount = 0 Then
            picTmpBrowser.Cls
            picBoardBrowser_Paint
      End If
      
      Exit Sub
ErrorHandler:
      On Error Resume Next
      Dir1.Path = File1.Path
End Sub


Sub DisplayLevel(ByVal FileName As String)
      On Error GoTo ErrorHandler
      
      If Trim(FileName) = "" Then Exit Sub
      
      Dim col As Integer, row As Integer
      Dim FileNum As Integer
      Dim GameData As String
      Dim ObjectType As Integer
            
      FileNum = FreeFile()
      
      'Reading in game data from a file
      Open FileName For Input As #FileNum
            For row = 0 To (BOARD_DIMENSION_Y - 1)
                  Input #FileNum, GameData
                              
                  For col = 0 To (BOARD_DIMENSION_X - 1)
                        'Store the game object types
                        tmpBoard(col, row) = Val(Mid(GameData, col + 1, 1))
                  Next col
            Next row
      Close #FileNum
      
            
      If Not CanLoadLevel(tmpBoard()) Then
            
            Dim AdjustedX As Integer, AdjustedY As Integer
            Dim strMsg As String
            Dim TextSize As Size
                  
            With frmLevels
                  .cmdOK.Enabled = False
                  
                  .picTmpBrowser.Cls
                  .picBoardBrowser.Cls
                  .picBoardBrowser.Font = "Arial"
                  .picBoardBrowser.FontBold = True
                  .picBoardBrowser.FontSize = "28"
                  
                  strMsg = "Invalid level file"
                  
                  Call GetTextExtentPoint(.picBoardBrowser.hdc, strMsg, Len(strMsg), TextSize)
                  
                  AdjustedX = (.picBoardBrowser.ScaleWidth - TextSize.cx) \ 2
                  AdjustedY = (.picBoardBrowser.ScaleHeight - TextSize.cy) \ 2
                  
                  .picBoardBrowser.CurrentX = AdjustedX
                  .picBoardBrowser.CurrentY = AdjustedY
                  .picBoardBrowser.Print "Invalid level file"
            End With
            
            Exit Sub
      End If
      
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  ObjectType = tmpBoard(col, row)
                              
                  BitBlt picTmpBrowser.hdc, _
                          col * OBJECT_WIDTH, row * OBJECT_HEIGHT, _
                          OBJECT_WIDTH, OBJECT_HEIGHT, _
                          picObjectsBrowser.hdc, ObjectType * OBJECT_WIDTH, 0, vbSrcCopy
            Next col
      Next row
      
      picBoardBrowser_Paint
      
      frmLevels.cmdOK.Enabled = True
      Exit Sub
ErrorHandler:
      frmLevels.cmdOK.Enabled = False
End Sub

Private Sub picBoardBrowser_Paint()
      BitBlt picBoardBrowser.hdc, 0, 0, BoardWidth, BoardHeight, picTmpBrowser.hdc, 0, 0, vbSrcCopy
End Sub
