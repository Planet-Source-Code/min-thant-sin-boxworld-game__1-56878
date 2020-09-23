Attribute VB_Name = "basDeclarations"
Option Explicit

Public Type COORD
      X As Integer
      Y As Integer
End Type

'Game board dimensions
Public Const BOARD_DIMENSION_X As Integer = 16       'There are 16 squares across...
Public Const BOARD_DIMENSION_Y As Integer = 14       'and 14 squares down

'Every object in the game is of the same size (in pixels)
Public Const OBJECT_WIDTH As Integer = 30
Public Const OBJECT_HEIGHT As Integer = 30

'Game objects (7 objects altogether)
Public Const GRAY_WALL As Integer = 0   'Not really a wall, only used for display
Public Const WHITE_WALL As Integer = 1   'Inner walls, cannot be passed through
Public Const BLUE_FLOOR As Integer = 2   'Floor on which the little boy walks around
Public Const YELLOW_BOX As Integer = 3   'The box which is not in the destination position
Public Const RED_BOX As Integer = 4         'The box which is in the destination position
Public Const LITTLE_BALL As Integer = 5    'Indicates destination position
Public Const LITTLE_BOY As Integer = 6     'The little boy which the player controls

Public BoardWidth As Integer      'Gameboard's width in pixel
Public BoardHeight As Integer     'Gameboard's height in pixel

Public strErrorMsg As String

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
