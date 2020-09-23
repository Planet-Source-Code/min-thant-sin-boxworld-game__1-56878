Attribute VB_Name = "basMain"
Option Explicit

Public Sub Main()
      Dim ExtraWidth As Integer, ExtraHeight As Integer
      
      'Gameboard's width and height in pixels
      BoardWidth = (BOARD_DIMENSION_X * OBJECT_WIDTH)
      BoardHeight = (BOARD_DIMENSION_Y * OBJECT_HEIGHT)
      
      NumObjects = (BOARD_DIMENSION_X * BOARD_DIMENSION_Y)
      
      ReDim Objects(0 To NumObjects - 1)
      ReDim Board(0 To BOARD_DIMENSION_X - 1, 0 To BOARD_DIMENSION_Y - 1)
      ReDim tmpBoard(0 To BOARD_DIMENSION_X - 1, 0 To BOARD_DIMENSION_Y - 1)
      ReDim UndoPoint(0 To BOARD_DIMENSION_X - 1, 0 To BOARD_DIMENSION_Y - 1)
      
      With frmMain.picBoard
            .Move 0, 0
            .Width = .ScaleX(BoardWidth, .ScaleMode, vbTwips)
            .Height = .ScaleY(BoardHeight, .ScaleMode, vbTwips)
      End With
      
      With frmMain.picTmp
            .Move 0, 0
            .Width = .ScaleX(BoardWidth, .ScaleMode, vbTwips)
            .Height = .ScaleY(BoardHeight, .ScaleMode, vbTwips)
      End With
      
      
      With frmMain
            ExtraWidth = .Width - .ScaleWidth
            ExtraHeight = .Height - .ScaleHeight
      
            .Move .Left, .Top, _
            .picBoard.Width + ExtraWidth, .picBoard.Height + ExtraHeight + 5
      End With
      
      With frmLevels.picBoardBrowser
            .Width = .ScaleX(BoardWidth, .ScaleMode, vbTwips)
            .Height = .ScaleY(BoardHeight, .ScaleMode, vbTwips)
      End With

      With frmLevels.picTmpBrowser
            .Width = .ScaleX(BoardWidth, .ScaleMode, vbTwips)
            .Height = .ScaleY(BoardHeight, .ScaleMode, vbTwips)
      End With

      On Error GoTo ErrorHandler
      
      If Dir(App.Path & "\Levels", vbDirectory) <> "" Then
            frmLevels.Dir1.Path = App.Path & "\Levels"
      Else
            frmLevels.Dir1.Path = App.Path
      End If
      
      If frmLevels.File1.ListCount = 0 Then
            MsgBox "No level files found in the current directory." & vbCrLf & _
                       "You will have to manually search for the level files", vbInformation, "BoxWorld"
      End If
      
      NumGameFiles = frmLevels.File1.ListCount
      
      GameLevel = GetSetting(MY_APP, MY_SECTION, MY_KEY, 0)
      
      If GameLevel >= NumGameFiles Then
            GameLevel = 0
      End If
      
      Load frmMain
      frmMain.Show
      frmMain.Refresh
      
      NumUndoPoints = 0
      
      If frmLevels.File1.ListCount > 0 Then
            frmLevels.File1.ListIndex = GameLevel
            LoadGame AddASlash(frmLevels.File1.Path) & frmLevels.File1.FileName
      End If
      
      Exit Sub
      
ErrorHandler:
      MsgBox Err.Description
      Call WriteErrorLog
      'End
End Sub
