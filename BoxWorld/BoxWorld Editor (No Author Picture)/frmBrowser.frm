VERSION 5.00
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BoxWorld Level Browser"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   360
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   4215
   End
   Begin VB.PictureBox picObjects 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   375
      Picture         =   "frmBrowser.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   3225
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   375
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   1
      Top             =   3825
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   7575
      TabIndex        =   7
      Top             =   6525
      Width           =   1965
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Open in Level Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4350
      TabIndex        =   6
      Top             =   6525
      Width           =   3090
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   9600
      TabIndex        =   5
      Top             =   6525
      Width           =   1965
   End
   Begin VB.FileListBox File1 
      Height          =   2250
      Left            =   75
      Pattern         =   "*.bxw"
      TabIndex        =   4
      Top             =   5025
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   4410
      Left            =   75
      TabIndex        =   3
      Top             =   525
      Width           =   4215
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6300
      Left            =   4350
      ScaleHeight     =   418
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   2
      Top             =   75
      Width           =   7200
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Board() As Integer           'Board(x,y) stores an object's type

Private Sub cmdAbout_Click()
      frmAbout.lblAbout(0) = "BoxWorld Level Browser"
      frmAbout.lblAbout(1) = "BoxWorld Level Browser"
      
      frmAbout.Show vbModal
End Sub

Private Sub cmdClose_Click()
      Unload Me
End Sub

Private Sub cmdLoad_Click()
      Me.Hide
      frmMain.Caption = "BoxWorld - " & File1.FileName
      Call frmMain.LoadLevel(AddASlash(File1.Path) & File1.FileName)
End Sub

Private Sub Dir1_Change()
      On Error GoTo ErrorHandler
      File1.Path = Dir1.Path
      
      Exit Sub
ErrorHandler:
      On Error Resume Next
      Dir1.Path = File1.Path
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
      Dim FileName As String
      
      FileName = AddASlash(File1.Path) & File1.FileName
      
      Call DisplayLevel(FileName)
      
      cmdLoad.Enabled = (File1.ListIndex > -1)
End Sub

Function AddASlash(ByVal strInput As String) As String
      AddASlash = strInput & "\"
      If Right$(strInput, 1) = "\" Then AddASlash = strInput
End Function

Private Sub picBoard_Paint()
      BitBlt picBoard.hDC, 0, 0, BoardWidth, BoardHeight, picTmp.hDC, 0, 0, vbSrcCopy
End Sub

Private Sub Form_Load()
      'Gameboard's width and height in pixels
      BoardWidth = (BOARD_DIMENSION_X * OBJECT_WIDTH)
      BoardHeight = (BOARD_DIMENSION_Y * OBJECT_HEIGHT)
      
      ReDim Board(0 To BOARD_DIMENSION_X - 1, 0 To BOARD_DIMENSION_Y - 1)
      
      With picBoard
            .Width = .ScaleX(BoardWidth, .ScaleMode, vbTwips)
            .Height = .ScaleY(BoardHeight, .ScaleMode, vbTwips)
      End With
      
      With picTmp
            .Width = .ScaleX(BoardWidth, .ScaleMode, vbTwips)
            .Height = .ScaleY(BoardHeight, .ScaleMode, vbTwips)
      End With
End Sub

Sub DisplayLevel(ByVal FileName As String)
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
                        Board(col, row) = Mid$(GameData, col + 1, 1)
                  Next col
            Next row
      Close #FileNum
      
      
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  ObjectType = Board(col, row)
                              
                  BitBlt picTmp.hDC, _
                          col * OBJECT_WIDTH, row * OBJECT_HEIGHT, _
                          OBJECT_WIDTH, OBJECT_HEIGHT, _
                          picObjects.hDC, ObjectType * OBJECT_WIDTH, 0, vbSrcCopy
            Next col
      Next row
      
      picBoard_Paint
      
End Sub

