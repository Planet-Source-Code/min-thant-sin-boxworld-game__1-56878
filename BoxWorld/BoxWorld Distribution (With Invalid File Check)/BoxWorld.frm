VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BoxWorld - Recreated by Min Thant Sin"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBoy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3000
      Picture         =   "BoxWorld.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   2625
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   1
      Top             =   1800
      Width           =   2190
   End
   Begin VB.PictureBox picObjects 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   300
      Picture         =   "BoxWorld.frx":2A72
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   2625
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   3
      Top             =   675
      Width           =   2190
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRetry 
         Caption         =   "&Retry"
         Shortcut        =   ^R
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToWorld 
         Caption         =   "Go to &world..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuLaunchEditor 
         Caption         =   "&Launch Level Editor..."
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHowToPlay 
         Caption         =   "&How to play..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If boolGameOver Then Exit Sub
      If boolBoyExists = False Then Exit Sub
      
      Dim LeftObject As Integer
      Dim RightObject As Integer
      Dim TopObject As Integer
      Dim BottomObject As Integer
      Dim i As Integer, TargetBox As Integer
      Dim AdjacentObject As Integer
      
      Select Case KeyCode
      Case vbKeyLeft
            LeftObject = GetObjectFromPosition(Boy.XPos - 1, Boy.YPos)
            
            If LeftObject = WHITE_WALL Or LeftObject = GRAY_WALL Then Exit Sub
            
            BoyDirection = DIRECTION_LEFT
            
            If LeftObject = YELLOW_BOX Or LeftObject = RED_BOX Then
                  For i = 0 To NumObjects - 1
                        If (Objects(i).XPos = Boy.XPos - 1) And (Objects(i).YPos = Boy.YPos) Then
                              If Objects(i).Type = YELLOW_BOX Or Objects(i).Type = RED_BOX Then
                                    TargetBox = i
                                    Exit For
                              End If
                        End If
                  Next i
                  
                  AdjacentObject = GetObjectFromPosition(Boy.XPos - 2, Boy.YPos)
                  If AdjacentObject = YELLOW_BOX Or _
                     AdjacentObject = RED_BOX Or _
                     AdjacentObject = WHITE_WALL Or _
                     AdjacentObject = GRAY_WALL Then Exit Sub
                   
                  Call CreateUndoPoint
                   
                  Call ClearBoyPosition
                          
                  If Objects(TargetBox).Type = RED_BOX Then
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = LITTLE_BALL
                  Else
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = BLUE_FLOOR
                  End If
                  
                  If Board(Objects(TargetBox).XPos - 1, Objects(TargetBox).YPos) = LITTLE_BALL Then
                        Objects(TargetBox).Type = RED_BOX
                        Board(Objects(TargetBox).XPos - 1, Objects(TargetBox).YPos) = RED_BOX
                  Else
                        Objects(TargetBox).Type = YELLOW_BOX
                        Board(Objects(TargetBox).XPos - 1, Objects(TargetBox).YPos) = YELLOW_BOX
                  End If
                  
                  'Update the little boy's and objects' data
                  Boy.XPos = Boy.XPos - 1
                  Boy.Left = Boy.XPos * OBJECT_WIDTH
                  
                  Objects(TargetBox).XPos = Objects(TargetBox).XPos - 1
                  Objects(TargetBox).Left = Objects(TargetBox).XPos * OBJECT_WIDTH
                  
                  'Display updated objects
                  BitBlt picBoard.hdc, _
                          Objects(TargetBox).Left, Objects(TargetBox).Top, _
                          OBJECT_WIDTH, OBJECT_HEIGHT, picObjects.hdc, _
                          Objects(TargetBox).Type * OBJECT_WIDTH, 0, vbSrcCopy
                          
                  BitBlt picBoard.hdc, _
                          Boy.Left, Boy.Top, OBJECT_WIDTH, OBJECT_HEIGHT, _
                          picBoy.hdc, BoyDirection * OBJECT_WIDTH, 0, vbSrcCopy
                  
            Else
                  Call ClearBoyPosition
            
                  Boy.XPos = Boy.XPos - 1
                  
                  If Boy.XPos < 0 Then Boy.XPos = 0
                  Boy.Left = Boy.XPos * OBJECT_WIDTH
            End If
            
            
      Case vbKeyRight
            RightObject = GetObjectFromPosition(Boy.XPos + 1, Boy.YPos)
            
            If RightObject = WHITE_WALL Or RightObject = GRAY_WALL Then Exit Sub
            
            BoyDirection = DIRECTION_RIGHT
                  
            If RightObject = YELLOW_BOX Or RightObject = RED_BOX Then
                  
                  For i = 0 To NumObjects - 1
                        If Objects(i).Type = YELLOW_BOX Or Objects(i).Type = RED_BOX Then
                              If (Objects(i).XPos = Boy.XPos + 1) And (Objects(i).YPos = Boy.YPos) Then
                                    TargetBox = i
                                    Exit For
                              End If
                        End If
                  Next i
                  
                  AdjacentObject = GetObjectFromPosition(Boy.XPos + 2, Boy.YPos)
                  If AdjacentObject = YELLOW_BOX Or _
                     AdjacentObject = RED_BOX Or _
                     AdjacentObject = WHITE_WALL Or _
                     AdjacentObject = GRAY_WALL Then Exit Sub
                     
                   
                  Call CreateUndoPoint
                   
                   
                  Call ClearBoyPosition
                  
                  If Objects(TargetBox).Type = RED_BOX Then
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = LITTLE_BALL
                  Else
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = BLUE_FLOOR
                  End If
                  
                  If Board(Objects(TargetBox).XPos + 1, Objects(TargetBox).YPos) = LITTLE_BALL Then
                        Objects(TargetBox).Type = RED_BOX
                        Board(Objects(TargetBox).XPos + 1, Objects(TargetBox).YPos) = RED_BOX
                  Else
                        Objects(TargetBox).Type = YELLOW_BOX
                        Board(Objects(TargetBox).XPos + 1, Objects(TargetBox).YPos) = YELLOW_BOX
                  End If
                  
                  Boy.XPos = Boy.XPos + 1
                  Boy.Left = Boy.XPos * OBJECT_WIDTH
                  
                  Objects(TargetBox).XPos = Objects(TargetBox).XPos + 1
                  Objects(TargetBox).Left = Objects(TargetBox).XPos * OBJECT_WIDTH
                  
                  BitBlt picBoard.hdc, _
                          Objects(TargetBox).Left, Objects(TargetBox).Top, _
                          OBJECT_WIDTH, OBJECT_HEIGHT, picObjects.hdc, _
                          Objects(TargetBox).Type * OBJECT_WIDTH, 0, vbSrcCopy
                          
                  BitBlt picBoard.hdc, _
                          Boy.Left, Boy.Top, OBJECT_WIDTH, OBJECT_HEIGHT, _
                          picBoy.hdc, BoyDirection * OBJECT_WIDTH, 0, vbSrcCopy
                  
            Else
                  Call ClearBoyPosition
                  
                  Boy.XPos = Boy.XPos + 1
                  If Boy.XPos > (BOARD_DIMENSION_X - 1) Then
                        Boy.XPos = (BOARD_DIMENSION_X - 1)
                  End If
                  Boy.Left = Boy.XPos * OBJECT_WIDTH
            End If
            
      Case vbKeyUp
            TopObject = GetObjectFromPosition(Boy.XPos, Boy.YPos - 1)
      
            If TopObject = WHITE_WALL Or TopObject = GRAY_WALL Then Exit Sub
            
            BoyDirection = DIRECTION_UP
                  
            If TopObject = YELLOW_BOX Or TopObject = RED_BOX Then
                  
                  For i = 0 To NumObjects - 1
                        If Objects(i).Type = YELLOW_BOX Or Objects(i).Type = RED_BOX Then
                              If (Objects(i).YPos = Boy.YPos - 1) And (Objects(i).XPos = Boy.XPos) Then
                                    TargetBox = i
                                    Exit For
                              End If
                        End If
                  Next i
                  
                  AdjacentObject = GetObjectFromPosition(Boy.XPos, Boy.YPos - 2)
                  If AdjacentObject = YELLOW_BOX Or _
                     AdjacentObject = RED_BOX Or _
                     AdjacentObject = WHITE_WALL Or _
                     AdjacentObject = GRAY_WALL Then Exit Sub
                   
                  Call CreateUndoPoint
                   
                   
                  Call ClearBoyPosition
                  
                  If Objects(TargetBox).Type = RED_BOX Then
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = LITTLE_BALL
                  Else
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = BLUE_FLOOR
                  End If
                  
                  If Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos - 1) = LITTLE_BALL Then
                        Objects(TargetBox).Type = RED_BOX
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos - 1) = RED_BOX
                  Else
                        Objects(TargetBox).Type = YELLOW_BOX
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos - 1) = YELLOW_BOX
                  End If
                  
                  Boy.YPos = Boy.YPos - 1
                  Boy.Top = Boy.YPos * OBJECT_HEIGHT
                  
                  Objects(TargetBox).YPos = Objects(TargetBox).YPos - 1
                  Objects(TargetBox).Top = Objects(TargetBox).YPos * OBJECT_HEIGHT
                  
                  BitBlt picBoard.hdc, _
                          Objects(TargetBox).Left, Objects(TargetBox).Top, _
                          OBJECT_WIDTH, OBJECT_HEIGHT, picObjects.hdc, _
                          Objects(TargetBox).Type * OBJECT_WIDTH, 0, vbSrcCopy
                          
                  BitBlt picBoard.hdc, _
                          Boy.Left, Boy.Top, OBJECT_WIDTH, OBJECT_HEIGHT, _
                          picBoy.hdc, BoyDirection * OBJECT_WIDTH, 0, vbSrcCopy
                  
            Else
                  Call ClearBoyPosition
                  
                  Boy.YPos = Boy.YPos - 1
                  
                  If Boy.YPos < 0 Then Boy.YPos = 0
                  Boy.Top = Boy.YPos * OBJECT_HEIGHT
            End If
            
            
      Case vbKeyDown
            BottomObject = GetObjectFromPosition(Boy.XPos, Boy.YPos + 1)
            
            If BottomObject = WHITE_WALL Or BottomObject = GRAY_WALL Then Exit Sub
            
            BoyDirection = DIRECTION_DOWN
            
            If BottomObject = YELLOW_BOX Or BottomObject = RED_BOX Then
                  
                  For i = 0 To NumObjects - 1
                        If Objects(i).Type = YELLOW_BOX Or Objects(i).Type = RED_BOX Then
                              If (Objects(i).YPos = Boy.YPos + 1) And (Objects(i).XPos = Boy.XPos) Then
                                    TargetBox = i
                                    Exit For
                              End If
                        End If
                  Next i
                  
                  AdjacentObject = GetObjectFromPosition(Boy.XPos, Boy.YPos + 2)
                  If AdjacentObject = YELLOW_BOX Or _
                     AdjacentObject = RED_BOX Or _
                     AdjacentObject = WHITE_WALL Or _
                     AdjacentObject = GRAY_WALL Then Exit Sub
                   
                  Call CreateUndoPoint
                   
                  Call ClearBoyPosition
                  
                  If Objects(TargetBox).Type = RED_BOX Then
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = LITTLE_BALL
                  Else
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos) = BLUE_FLOOR
                  End If
                  
                  If Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos + 1) = LITTLE_BALL Then
                        Objects(TargetBox).Type = RED_BOX
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos + 1) = RED_BOX
                  Else
                        Objects(TargetBox).Type = YELLOW_BOX
                        Board(Objects(TargetBox).XPos, Objects(TargetBox).YPos + 1) = YELLOW_BOX
                  End If
                  
                  Boy.YPos = Boy.YPos + 1
                  Boy.Top = Boy.YPos * OBJECT_HEIGHT
                  
                  Objects(TargetBox).YPos = Objects(TargetBox).YPos + 1
                  Objects(TargetBox).Top = Objects(TargetBox).YPos * OBJECT_HEIGHT
                  
                  BitBlt picBoard.hdc, _
                          Objects(TargetBox).Left, Objects(TargetBox).Top, _
                          OBJECT_WIDTH, OBJECT_HEIGHT, picObjects.hdc, _
                          Objects(TargetBox).Type * OBJECT_WIDTH, 0, vbSrcCopy
                          
                  BitBlt picBoard.hdc, _
                          Boy.Left, Boy.Top, OBJECT_WIDTH, OBJECT_HEIGHT, _
                          picBoy.hdc, BoyDirection * OBJECT_WIDTH, 0, vbSrcCopy
                  
            Else
                  Call ClearBoyPosition
                  
                  Boy.YPos = Boy.YPos + 1
                  
                  If Boy.YPos > (BOARD_DIMENSION_Y - 1) Then
                        Boy.YPos = (BOARD_DIMENSION_Y - 1)
                  End If
                  
                  Boy.Top = Boy.YPos * OBJECT_HEIGHT
            End If
            
      End Select
      
      Call UpdateBoyPosition
      
      BitBlt picTmp.hdc, 0, 0, BoardWidth, BoardHeight, picBoard.hdc, 0, 0, vbSrcCopy
      
      
      If PuzzleSolved() Then
            Beep
            boolGameOver = True
                                    
            GameLevel = GameLevel + 1
            
            'If the user has finished last game level...
            If GameLevel > (frmLevels.File1.ListCount - 1) Then
                  GameLevel = GameLevel - 1
                  frmFinished.Show vbModal
            Else
                  frmSolved.Show vbModal
            End If
            
      End If
End Sub

Sub ClearBoyPosition()
      On Error GoTo ErrorHandler
      
      BitBlt picBoard.hdc, Boy.Left, Boy.Top, OBJECT_WIDTH, OBJECT_HEIGHT, _
              picObjects.hdc, Board(Boy.XPos, Boy.YPos) * OBJECT_WIDTH, 0, vbSrcCopy
              
      Exit Sub
ErrorHandler:
      MsgBox Err.Description
End Sub

Sub UpdateBoyPosition()
      On Error GoTo ErrorHandler
      
      BitBlt picBoard.hdc, Boy.Left, Boy.Top, OBJECT_WIDTH, OBJECT_HEIGHT, _
              picBoy.hdc, BoyDirection * OBJECT_WIDTH, 0, vbSrcCopy
      
      Exit Sub
ErrorHandler:
      MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
      End
End Sub

Private Sub mnuAbout_Click()
      frmAbout.Show vbModal
End Sub

Private Sub mnuGame_Click()
      mnuLaunchEditor.Enabled = False
      mnuLaunchEditor.Caption = "&Launch Level Editor... (file not found)"
      
      If UCase(Dir(App.Path & "\BxwEdit.exe")) = UCase("BxwEdit.exe") Then
            mnuLaunchEditor.Enabled = True
            mnuLaunchEditor.Caption = "&Launch Level Editor..."
      End If
      
      mnuUndo.Enabled = (NumUndoPoints > 0)
End Sub

Private Sub mnuGoToWorld_Click()
      frmLevels.Show vbModal
End Sub

Private Sub mnuHowToPlay_Click()
      frmHowTo.Show vbModal
End Sub

Private Sub mnuLaunchEditor_Click()
      On Error GoTo ErrorHandler
      
      Shell App.Path & "\BxwEdit.exe", vbNormalFocus
      Exit Sub
ErrorHandler:
      MsgBox Err.Description, vbExclamation, "BoxWorld"
End Sub

Private Sub mnuQuit_Click()
      End
End Sub

Private Sub mnuRetry_Click()
      frmLevels.File1.ListIndex = GameLevel
      LoadGame AddASlash(frmLevels.File1.Path) & frmLevels.File1.FileName
End Sub

Private Sub mnuUndo_Click()
      If NumUndoPoints = 0 Then Exit Sub
      
      Call UndoMovement
End Sub

Private Sub picBoard_Paint()
      BitBlt picBoard.hdc, 0, 0, BoardWidth, BoardHeight, picTmp.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub CreateUndoPoint()
      Dim col As Integer, row As Integer
      
      If NumUndoPoints = 0 Then
            NumUndoPoints = NumUndoPoints + 1
      End If
      
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  UndoPoint(col, row) = Board(col, row)
            Next col
      Next row
      
      tmpBoy = Boy
End Sub

Public Sub UndoMovement()
      Dim i As Integer
      Dim col As Integer, row As Integer
      Dim ObjectType As Integer
      
      i = 0
            
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  
                  ObjectType = UndoPoint(col, row)
                  Board(col, row) = ObjectType
                  
                  With Objects(i)
                        .Type = ObjectType
                        .XPos = col
                        .YPos = row
                        .Left = col * OBJECT_WIDTH
                        .Top = row * OBJECT_HEIGHT
                        
                        BitBlt frmMain.picBoard.hdc, .Left, .Top, OBJECT_WIDTH, OBJECT_HEIGHT, frmMain.picObjects.hdc, .Type * OBJECT_WIDTH, 0, vbSrcCopy
                  End With
                  
                  i = i + 1
            Next col
      Next row
      
      Boy = tmpBoy
      
      BitBlt frmMain.picBoard.hdc, _
              Boy.Left, Boy.Top, _
              OBJECT_WIDTH, OBJECT_HEIGHT, frmMain.picBoy.hdc, _
              DIRECTION_DOWN * OBJECT_WIDTH, 0, vbSrcCopy
                  
      BitBlt frmMain.picTmp.hdc, 0, 0, BoardWidth, BoardHeight, _
              frmMain.picBoard.hdc, 0, 0, vbSrcCopy
            
      picBoard_Paint
            
      NumUndoPoints = NumUndoPoints - 1
End Sub
