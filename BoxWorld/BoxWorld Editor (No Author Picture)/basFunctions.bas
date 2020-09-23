Attribute VB_Name = "basFunctions"
Option Explicit

'Get coordinates under mouse position
Public Function GetCoord(ByVal X As Single, ByVal Y As Single) As COORD
      '-1 to indicate mouse position is out of Board range.
      
      GetCoord.X = -1
      GetCoord.Y = -1
      
      'Make sure mouse position is within range
      If X >= 0 And X <= frmMain.picBoard.ScaleWidth Then
            If Y >= 0 And Y <= frmMain.picBoard.ScaleHeight Then
            
                  'Calculate X & Y coordinates
                  GetCoord.X = Int(X / OBJECT_WIDTH)
                  GetCoord.Y = Int(Y / OBJECT_HEIGHT)
            End If
      End If
End Function
