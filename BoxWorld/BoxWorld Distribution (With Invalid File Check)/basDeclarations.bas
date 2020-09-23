Attribute VB_Name = "basDeclarations"
Option Explicit

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

Public Const DIRECTION_LEFT As Integer = 0      'The little boy is going to left direction
Public Const DIRECTION_RIGHT As Integer = 1    'The little boy is going to right direction
Public Const DIRECTION_UP As Integer = 2         'The little boy is going to upward direction
Public Const DIRECTION_DOWN As Integer = 3    'The little boy is going to downward direction

Public Const MY_APP  As String = "Min Thant Sin BoxWorld Application"
Public Const MY_SECTION As String = "Levels Section"
Public Const MY_KEY As String = "Level Reached"

Public Type GAME_OBJECT_DATA
      XPos As Integer    'Coordinate X position of the object
      YPos As Integer    'Coordinate Y position of the object
      Left As Integer     'Actual left position of the object on the gameboard
      Top As Integer     'Actual top position of the object on the gameboard
      Type As Integer   'The object's type (one of the seven types)
End Type

'////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Type DESTINATION_DATA
      XPos As Integer
      YPos As Integer
End Type

Public Type Size
      cx As Long
      cy As Long
End Type

Public Destinations() As DESTINATION_DATA
'////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Boy As GAME_OBJECT_DATA           'The little boy

Public Objects() As GAME_OBJECT_DATA    'All the other objects on the gameboard

Public Board() As Integer           'Board(x,y) stores an object's type
Public tmpBoard() As Integer      'Board(x,y) stores an object's type (for frmLevels)

Public GameLevel As Integer
Public NumGameFiles As Integer
Public NumObjects As Integer     'Number of objects on the gameboard
Public NumBoxesToMove As Integer
Public BoyDirection As Integer     'The direction in which the little boy is going

Public BoardWidth As Integer      'Gameboard's width in pixel
Public BoardHeight As Integer     'Gameboard's height in pixel

Public boolGameOver As Boolean
Public boolBoyExists As Boolean
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public NumUndoPoints As Integer

Public tmpBoy As GAME_OBJECT_DATA
Public UndoPoint() As Integer

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public strErrorMsg As String

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTextExtentPoint Lib "gdi32.dll" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
