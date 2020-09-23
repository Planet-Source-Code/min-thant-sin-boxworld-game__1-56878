VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BoxWorld Level Editor"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3375
      Top             =   7125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   150
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   244
      TabIndex        =   1
      Top             =   7650
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.PictureBox picObjects 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   150
      Picture         =   "BoxWorld.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   7125
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.PictureBox picEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   0
      ScaleHeight     =   6465
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.PictureBox picBoard 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6300
         Left            =   1725
         ScaleHeight     =   418
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   478
         TabIndex        =   10
         Top             =   75
         Width           =   7200
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "Gray Wall"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   0
         Left            =   75
         Picture         =   "BoxWorld.frx":4A52
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Used as walls surrounding white walls"
         Top             =   75
         Width           =   1590
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "White Wall"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   1
         Left            =   75
         Picture         =   "BoxWorld.frx":555C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "The inner walls which enclose the little boy and all the other objects"
         Top             =   975
         Width           =   1590
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "Blue Floor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   2
         Left            =   75
         Picture         =   "BoxWorld.frx":6066
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Blue floors on which the little boy walks around"
         Top             =   1875
         Width           =   1590
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "Yellow Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   3
         Left            =   75
         Picture         =   "BoxWorld.frx":6B70
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "The box which is not in the destination position"
         Top             =   2775
         Width           =   1590
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "Red Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   4
         Left            =   75
         Picture         =   "BoxWorld.frx":767A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "The box that is already in the destination position"
         Top             =   3675
         Width           =   1590
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "Little Ball"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   5
         Left            =   75
         Picture         =   "BoxWorld.frx":8184
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Balls marking the destination positions"
         Top             =   4575
         Width           =   1590
      End
      Begin VB.OptionButton optObjects 
         Caption         =   "Little Boy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   6
         Left            =   75
         Picture         =   "BoxWorld.frx":8C8E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "The little boy which the user controls to move the boxes"
         Top             =   5475
         Width           =   1590
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLevelBrowser 
         Caption         =   "Level &Browser..."
         Shortcut        =   ^B
      End
      Begin VB.Menu sepLevelBrowser 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadLevel 
         Caption         =   "&Load Level..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSaveLevel 
         Caption         =   "&Save Level..."
         Shortcut        =   ^S
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuMakeAll 
         Caption         =   "Make all objects &Gray Wall"
         Index           =   0
      End
      Begin VB.Menu mnuMakeAll 
         Caption         =   "Make all objects &White Wall"
         Index           =   1
      End
      Begin VB.Menu mnuMakeAll 
         Caption         =   "Make all objects &Blue Floor"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Board() As Integer           'Board(x,y) stores an object's type
Private CurrentObject As Integer

Private Sub Form_Unload(Cancel As Integer)
      End
End Sub

Private Sub mnuLevelBrowser_Click()
      frmBrowser.Show vbModal
End Sub

Private Sub mnuLoadLevel_Click()
      
      With CommonDialog1
            .FileName = ""
            .Filter = "BoxWorld Game File (*.bxw)|*.bxw"
            .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist
            .ShowOpen
      End With
            
      If CommonDialog1.FileName = "" Then Exit Sub
      
      frmLoading.lblMessage = "Loading level, please wait..."
      frmLoading.tmrLoading.Enabled = True
      frmLoading.Show vbModal
                  
      LoadLevel CommonDialog1.FileName
End Sub

'Make all objects on the gameboard either White Wall, Gray Wall, or Blue Floor
Private Sub mnuMakeAll_Click(Index As Integer)
      Call FillGameBoardWithObject(Index)
      Call UpdateGameBoardDisplay
End Sub

Private Sub mnuQuit_Click()
      End
End Sub

Private Sub mnuAbout_Click()
      frmAbout.lblAbout(0) = "BoxWorld Level Editor"
      frmAbout.lblAbout(1) = "BoxWorld Level Editor"
      frmAbout.Show vbModal
End Sub

Private Sub FillGameBoardWithObject(ByVal ObjectType As Integer)
      Dim row As Integer, col As Integer
      
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  Board(col, row) = ObjectType
            Next col
      Next row
      
End Sub

Private Sub mnuSaveLevel_Click()
      If Not CanSaveLevel() Then
            frmErrors.lblErrorMsg = strErrorMsg
            frmErrors.Show vbModal
            Exit Sub
      End If
      
      On Error GoTo ErrorHandler
      
      With CommonDialog1
            .FileName = ""
            .Filter = "BoxWorld Game File (*.bxw)|*.bxw"
            .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
            .ShowSave
      End With
      
      If CommonDialog1.FileName = "" Then Exit Sub
      
      frmLoading.lblMessage = "Saving level, please wait..."
      frmLoading.tmrLoading.Enabled = True
      frmLoading.Show vbModal
      
      SaveLevel CommonDialog1.FileName
      Exit Sub
ErrorHandler:
      MsgBox Err.Description
      'Do nothing
End Sub

Private Sub optObjects_Click(Index As Integer)
      CurrentObject = Index
End Sub

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
      
      picEditor.BackColor = Me.BackColor
      
      CurrentObject = GRAY_WALL
      optObjects(GRAY_WALL).Value = True
      
      Call FillGameBoardWithObject(CurrentObject)
      Call UpdateGameBoardDisplay
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      Dim XCoord As Integer, YCoord As Integer
      
      XCoord = GetCoord(X, Y).X
      YCoord = GetCoord(X, Y).Y
                  
      If XCoord = -1 Then Exit Sub
      If YCoord = -1 Then Exit Sub
      
      Board(XCoord, YCoord) = CurrentObject
      
      BitBlt picTmp.hDC, _
              XCoord * OBJECT_WIDTH, YCoord * OBJECT_HEIGHT, _
              OBJECT_WIDTH, OBJECT_HEIGHT, _
              picObjects.hDC, CurrentObject * OBJECT_WIDTH, 0, vbSrcCopy
             
      picBoard_Paint
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button = vbLeftButton Then
            Call picBoard_MouseDown(Button, Shift, X, Y)
      End If
End Sub

Public Sub LoadLevel(ByVal FileName As String)
      Dim col As Integer, row As Integer
      Dim FileNum As Integer
      Dim GameData As String
      Dim ObjectType As Integer
            
      On Error GoTo ErrorHandler
      
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
      
      Call UpdateGameBoardDisplay
      
      Exit Sub
ErrorHandler:
      
End Sub

Sub SaveLevel(ByVal FileName As String)
      Dim col As Integer, row As Integer
      Dim FileNum As Integer
      Dim GameData As String
      
      FileNum = FreeFile()
      Open FileName For Output As #FileNum
            
            For row = 0 To (BOARD_DIMENSION_Y - 1)
                  For col = 0 To (BOARD_DIMENSION_X - 1)
                        GameData = GameData & Board(col, row)
                  Next col
                        Print #FileNum, GameData
                        GameData = ""
            Next row
            
      Close #FileNum
End Sub

Sub UpdateGameBoardDisplay()
      Dim row As Integer, col As Integer
      Dim ObjectType As Integer
      
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

Function CanSaveLevel() As Boolean
      Dim col As Integer, row As Integer
      Dim NumBalls As Integer
      Dim NumYellowBoxes As Integer
      Dim NumRedBoxes As Integer
      Dim NumBoys As Integer
      Dim NumErrors As Integer
      Dim ObjectType As Integer
      
      
      NumBalls = 0
      NumYellowBoxes = 0
      NumRedBoxes = 0
      NumBoys = 0
      NumErrors = 0
      strErrorMsg = ""
      
      CanSaveLevel = True
      
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  ObjectType = Board(col, row)
                  
                  Select Case ObjectType
                  Case LITTLE_BALL
                        NumBalls = NumBalls + 1
                  Case YELLOW_BOX
                        NumYellowBoxes = NumYellowBoxes + 1
                        
                  'This code fragment is not necessary
                  '//////////////////////////////////////////////////////////////////
                  Case RED_BOX
                        NumRedBoxes = NumRedBoxes + 1
                  '//////////////////////////////////////////////////////////////////
                  
                  Case LITTLE_BOY
                        NumBoys = NumBoys + 1
                  End Select
            Next col
      Next row
      
      
      If NumBoys = 0 Or NumBoys > 1 Then
            NumErrors = NumErrors + 1
            
            If NumBoys = 0 Then
                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description: No (little boy) found on the gameboard." & vbCrLf
                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Place one (little boy) on the gameboard." & vbCrLf
                  strErrorMsg = strErrorMsg & vbNewLine
            Else
                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description: Too many (little boys) on the gameboard." & vbCrLf
                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Remove unnecessary (little boys) from the gameboard." & vbCrLf
                  strErrorMsg = strErrorMsg & vbNewLine
            End If
            
            CanSaveLevel = False
      End If
      
      'Check for "No scrambled boxes"
      If NumYellowBoxes = 0 Then
            NumErrors = NumErrors + 1
            
            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : No (yellow boxes) found on the gameboard." & vbCrLf
            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Place a certain number of (yellow boxes) on the gameboard." & vbCrLf
            strErrorMsg = strErrorMsg & vbNewLine
            
            CanSaveLevel = False
      End If
      
      'Check for "No destination positions" (no little balls)
      If NumBalls = 0 Then
            NumErrors = NumErrors + 1
            
            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : No (little balls) found on the gameboard." & vbCrLf
            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Place a certain number of (little balls) on the gameboard." & vbCrLf
            strErrorMsg = strErrorMsg & vbNewLine
            
            CanSaveLevel = False
      End If
      
      If NumBalls > NumYellowBoxes Then
            If NumYellowBoxes > 0 Then
                  NumErrors = NumErrors + 1

                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : The number of (little balls) is greater than the number of BOXES." & vbCrLf
                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Reduce the number of (little balls) on the gameboard." & vbCrLf
                  strErrorMsg = strErrorMsg & vbNewLine
            
                  CanSaveLevel = False
            End If
      Else
      
            If NumBalls < NumYellowBoxes Then
                  If NumBalls > 0 Then
                        NumErrors = NumErrors + 1
                  
                        strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : The number of (little balls) is less than the number of BOXES." & vbCrLf
                        strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Increase the number of (little balls) on the gameboard." & vbCrLf
                        strErrorMsg = strErrorMsg & vbNewLine
                  
                        CanSaveLevel = False
                  End If
            End If
      End If
End Function
