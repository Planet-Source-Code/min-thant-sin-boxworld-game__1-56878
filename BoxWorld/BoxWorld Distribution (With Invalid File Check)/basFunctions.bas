Attribute VB_Name = "basFunctions"
Option Explicit

Public Sub SaveGame()
      SaveSetting MY_APP, MY_SECTION, MY_KEY, CStr(GameLevel)
End Sub

Public Function AddASlash(ByVal strIn As String) As String
      If Right$(strIn, 1) = "\" Then
            AddASlash = strIn
      Else
            AddASlash = strIn & "\"
      End If
      
End Function

Public Function GetObjectFromPosition(ByVal X As Integer, ByVal Y As Integer) As Integer
      GetObjectFromPosition = -1
      
      If X >= 0 And X <= (BOARD_DIMENSION_X - 1) Then
            If Y >= 0 And Y <= (BOARD_DIMENSION_Y - 1) Then
                  GetObjectFromPosition = Board(X, Y)
            End If
      End If
End Function

Public Function PuzzleSolved() As Boolean
      On Error GoTo ErrorHandler
      
      Dim i As Integer, j As Integer
      
      PuzzleSolved = False
      
      j = 0
      For i = 0 To UBound(Destinations())
            If Board(Destinations(i).XPos, Destinations(i).YPos) = RED_BOX Then
                  j = j + 1
                  If j >= NumBoxesToMove Then
                        PuzzleSolved = True
                        Exit Function
                  End If
            End If
      Next i
      
      Exit Function
ErrorHandler:
      MsgBox Err.Description
      'Do nothing
End Function

Public Sub LoadGame(ByVal FileName As String)
      Dim i As Integer
      Dim row As Integer
      Dim col As Integer
      Dim ObjectType As Integer
      Dim GameData As String
            
      On Error GoTo ErrorHandler
      
      'Be sure to reset some variables
      '/////////////////////////////////////////////////////////////////////////////////////////////////////////
      NumBoxesToMove = 0
      boolBoyExists = False
      '/////////////////////////////////////////////////////////////////////////////////////////////////////////
      
      
      'Reading in game data from a file
      Open FileName For Input As #1
            For row = 0 To (BOARD_DIMENSION_Y - 1)
                  Input #1, GameData
                              
                  For col = 0 To (BOARD_DIMENSION_X - 1)
                        'Store the game object types
                        Board(col, row) = Mid(GameData, col + 1, 1)
                        
                        If Board(col, row) = LITTLE_BOY Then
                              boolBoyExists = True
                        End If
                        
                  Next col
            Next row
      Close #1
      
       If Not CanLoadLevel(Board()) Then
            boolGameOver = True
            
            frmMain.picTmp.Cls
            frmMain.picBoard.Cls
            
            MsgBox "Error loading level file.", vbExclamation, "BoxWorld"
            
            Exit Sub
      End If
      
      'Fill in game objects data
      i = 0
            
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  ObjectType = Board(col, row)
                  
                  With Objects(i)
                        .Type = ObjectType
                        .XPos = col
                        .YPos = row
                        .Left = col * OBJECT_WIDTH
                        .Top = row * OBJECT_HEIGHT
                        
                        BitBlt frmMain.picBoard.hdc, .Left, .Top, OBJECT_WIDTH, OBJECT_HEIGHT, frmMain.picObjects.hdc, .Type * OBJECT_WIDTH, 0, vbSrcCopy
                  End With
                  
                  
                  If ObjectType = LITTLE_BOY Then
                        Boy = Objects(i)
                        Board(col, row) = BLUE_FLOOR
                        
                        BitBlt frmMain.picBoard.hdc, _
                                Objects(i).Left, Objects(i).Top, _
                                OBJECT_WIDTH, OBJECT_HEIGHT, frmMain.picBoy.hdc, _
                                DIRECTION_DOWN * OBJECT_WIDTH, 0, vbSrcCopy
                  End If
                                                            
                                                            
                  'Store destination positions
                  '////////////////////////////////////////////////////////////////////////////////////////////////////////
                  If ObjectType = LITTLE_BALL Or _
                     ObjectType = RED_BOX Then
                        NumBoxesToMove = NumBoxesToMove + 1
                        
                        ReDim Preserve Destinations(0 To NumBoxesToMove - 1)
                        With Destinations(NumBoxesToMove - 1)
                              .XPos = col
                              .YPos = row
                        End With
                  End If
                  '////////////////////////////////////////////////////////////////////////////////////////////////////////
                  
                  i = i + 1
            Next col
      Next row
      
      BitBlt frmMain.picTmp.hdc, 0, 0, BoardWidth, BoardHeight, _
              frmMain.picBoard.hdc, 0, 0, vbSrcCopy
            
      'Indicate that the game has started
      boolGameOver = False
      
      Exit Sub
ErrorHandler:
      boolGameOver = True
      frmMain.picTmp.Cls
      frmMain.picBoard.Refresh
      MsgBox Err.Description
      Call WriteErrorLog
End Sub

Public Function CanLoadLevel(ByRef BoardData() As Integer) As Boolean
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
      
      CanLoadLevel = True
      
      For row = 0 To (BOARD_DIMENSION_Y - 1)
            For col = 0 To (BOARD_DIMENSION_X - 1)
                  ObjectType = BoardData(col, row)
                  
                  Select Case ObjectType
                  Case LITTLE_BALL
                        NumBalls = NumBalls + 1
                  Case YELLOW_BOX
                        NumYellowBoxes = NumYellowBoxes + 1
                  Case LITTLE_BOY
                        NumBoys = NumBoys + 1
                  End Select
            Next col
      Next row
      
      
      If NumBoys = 0 Or NumBoys > 1 Then
'            NumErrors = NumErrors + 1
'
'            If NumBoys = 0 Then
'                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description: No (little boy) found on the gameboard." & vbCrLf
'                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Place one (little boy) on the gameboard." & vbCrLf
'                  strErrorMsg = strErrorMsg & vbNewLine
'            Else
'                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description: Too many (little boys) on the gameboard." & vbCrLf
'                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Remove unnecessary (little boys) from the gameboard." & vbCrLf
'                  strErrorMsg = strErrorMsg & vbNewLine
'            End If
'
            CanLoadLevel = False
      End If
      
      'Check for "No scrambled boxes"
      If NumYellowBoxes = 0 Then
'            NumErrors = NumErrors + 1
'
'            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : No (yellow boxes) found on the gameboard." & vbCrLf
'            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Place a certain number of (yellow boxes) on the gameboard." & vbCrLf
'            strErrorMsg = strErrorMsg & vbNewLine
            
            CanLoadLevel = False
      End If
      
      'Check for "No destination positions" (no little balls)
      If NumBalls = 0 Then
'            NumErrors = NumErrors + 1
'
'            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : No (little balls) found on the gameboard." & vbCrLf
'            strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Place a certain number of (little balls) on the gameboard." & vbCrLf
'            strErrorMsg = strErrorMsg & vbNewLine

            CanLoadLevel = False
      End If
      
      If NumBalls > NumYellowBoxes Then
            If NumYellowBoxes > 0 Then
'                  NumErrors = NumErrors + 1
'
'                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : The number of (little balls) is greater than the number of BOXES." & vbCrLf
'                  strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Reduce the number of (little balls) on the gameboard." & vbCrLf
'                  strErrorMsg = strErrorMsg & vbNewLine
            
                  CanLoadLevel = False
            End If
      Else
      
            If NumBalls < NumYellowBoxes Then
                  If NumBalls > 0 Then
'                        NumErrors = NumErrors + 1
'
'                        strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Error Description : The number of (little balls) is less than the number of BOXES." & vbCrLf
'                        strErrorMsg = strErrorMsg & "(" & NumErrors & ") " & "Possible Solution : Increase the number of (little balls) on the gameboard." & vbCrLf
'                        strErrorMsg = strErrorMsg & vbNewLine
                  
                        CanLoadLevel = False
                  End If
            End If
      End If
End Function

Public Sub WriteErrorLog()
      On Error Resume Next
      
      Open App.Path & "\Errors.log" For Append As #1
            Print #1, "[Log Started]"
            Print #1, "***********************************"
            Print #1, "Error Number : " & Err.Number
            Print #1, "Error Source : " & Err.Source
            Print #1, "Error Description : " & Err.Description
            Print #1, "Date of Error Occurred : " & Format$(Date, "dd dddd mmmm yyyy")
            Print #1, "Time of Error Occurred : " & Time
            Print #1, "***********************************"
            Print #1, "[Log Ended]"
            Print #1, ""
      Close #1
End Sub
