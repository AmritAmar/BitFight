Attribute VB_Name = "DeclarationModule"
'-----------------------------------------------------------------
'                       DECLARATIONS
'  All Declarations in the whole Program happend in this Module
'-----------------------------------------------------------------

' --- USER BIT DECLARATIONS ---

Public UserBitX As Integer 'X Position of User Bit
Public UserBitY As Integer 'Y Position of User Bit

Public UserBitXSpeed As Integer 'Speed X of User Bit
Public UserBitYSpeed As Integer 'Speed Y of User Bit

Public UserBitUpSpeed As Integer 'Directional Speed of User Bit (UP)
Public UserBitDownSpeed As Integer 'Directional Speed of User Bit (DDOWN)
Public UserBitLeftSpeed As Integer 'Directional Speed of User Bit (LEFT)
Public UserBitRightSpeed As Integer 'Directional Speed of User Bit (RIGHT)

Public SpeedIncrementation As Integer 'Speed Increment
Public TopFire As Boolean 'Is Top Fire On
Public DownFire As Boolean
Public LeftFire As Boolean
Public RightFire As Boolean

Public Const UserBitXStart As Integer = 6635 'FIXED Starting Position X of User Bit
Public Const UserBitYStart As Integer = 2925 'FIXED Starting Position Y of User Bit

Public DistortionBitX As Integer 'X Position of Distortion Bit
Public DistortionBitY As Integer 'Y Position of Distortion Bit

' --- WALL DECLARATIONS ---

Public W(1 To 8) As Integer 'Array of 8 Wall Values

Public TopSafetyCounterX As Integer 'X Position of Top Safety Measure
Public TopSafetyCounterY As Integer 'Y Position of Top Safety Measure
Public LeftSafetyCounterX As Integer 'X Position of Left Safety Measure
Public LeftSafetyCounterY As Integer 'Y Position of Left Safety Measure
Public RightSafetyCounterX As Integer 'X Position of Right Safety Measure
Public RightSafetyCounterY As Integer 'Y Position of Right Safety Measure
Public BottomSafetyCounterX As Integer 'X Position of Bottom Safety Measure
Public BottomSafetyCounterY As Integer 'Y Position of Top Safety Measure

Public ZomBitEndWallX As Integer 'X Position of End Wall
Public ZomBitEndWallY As Integer 'Y Position of End Wall

' --- ENEMY BIT COLLISIONS ---

Public EnemyBit1X As Integer 'X Position of Enemy Bit 1
Public EnemyBit1Y As Integer 'Y Position of Enemy Bit 1
Public EnemyBit2X As Integer 'X Position of Enemy Bit 2
Public EnemyBit2Y As Integer 'Y Position of Enemy Bit 2
Public EnemyBit3X As Integer 'X Position of Enemy Bit 3
Public EnemyBit3Y As Integer 'Y Position of Enemy Bit 3
Public EnemyBit4X As Integer 'X Position of Enemy Bit 4
Public EnemyBit4Y As Integer 'Y Position of Enemy Bit 4
Public EnemyBit5X As Integer 'X Position of Enemy Bit 5
Public EnemyBit5Y As Integer 'Y Position of Enemy Bit 5
Public EnemyBit6X As Integer 'X Position of Enemy Bit 6
Public EnemyBit6Y As Integer 'Y Position of Enemy Bit 6

Public WarpMineBit1X As Integer 'X Position of Warp Mine 1
Public WarpMineBit1Y As Integer 'Y Position of Warp Mine 1
Public WarpMineBit2X As Integer 'X Position of Warp Mine 2
Public WarpMineBit2Y As Integer 'Y Position of Warp Mine 2

Public EnemyBit1XSpeed As Integer 'X Speed of Enemy Bit 1
Public EnemyBit1YSpeed As Integer 'Y Speed of Enemy Bit 1
Public EnemyBit2XSpeed As Integer 'X Speed of Enemy Bit 2
Public EnemyBit2YSpeed As Integer 'Y Speed of Enemy Bit 2
Public EnemyBit3XSpeed As Integer 'X Speed of Enemy Bit 3
Public EnemyBit3YSpeed As Integer 'Y Speed of Enemy Bit 3
Public EnemyBit4XSpeed As Integer 'X Speed of Enemy Bit 4
Public EnemyBit4YSpeed As Integer 'Y Speed of Enemy Bit 4
Public EnemyBit5XSpeed As Integer 'X Speed of Enemy Bit 5
Public EnemyBit5YSpeed As Integer 'Y Speed of Enemy Bit 5
Public EnemyBit6XSpeed As Integer 'X Speed of Enemy Bit 6
Public EnemyBit6YSpeed As Integer 'Y Speed of Enemy Bit 6

Public WarpMineBit1XSpeed As Integer 'X Speed of Warp Mine 1
Public WarpMineBit1YSpeed As Integer 'Y Speed of Warp Mine 1
Public WarpMineBit2XSpeed As Integer 'X Speed of Warp Mine 2
Public WarpMineBit2YSpeed As Integer 'Y Speed of Warp Mine 2

' --- POWER UP DECLARATIONS ---

Public PowerUpBitX As Integer 'X Position of Power Up
Public PowerUpBitY As Integer 'Y Position of Power Up

Public PowerUpPickUp As Boolean 'Checks if Power Up was taken
Public PowerUpBitNumber As Integer 'Random Integer for Power Up

' --- GAMEPLAY ELEMENTS DECLARATIONS ---

'Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long 'Beep Function
'Public BeepFreq As String 'Frequency of Beep
'Public BeepTime As String 'Time of Beep

Public isEnemyHit As Boolean 'True if User Gets hit
Public isSlowMotion As Boolean 'True if Slow Motion is On
Public isDistortion As Boolean 'true if Distortion is On
Public isEnemyDistortion As Boolean 'True if Enemy Collides with player when being Distorted
Public isBurntOut As Boolean 'True if User is too fast
Public isWall As Boolean 'True if Wall has been modified
Public isPlayerTP As Boolean 'True if Player is going to get teleported
Public isEMPOn As Boolean 'True if EMP is On and Charging
Public isEMPTrigerred As Boolean 'True if EMP will work
Public isWarpMineProximal As Boolean 'True if warp mine is too close to player
Public isZomBitsReAlign As Boolean 'True if ZomBits are Re-Aligning
Public isSolidColours As Boolean 'True if Solid Colours are Own
Public isExploded As Boolean 'True if User is Dead
Public isCircularPowerUp As Boolean 'True if Circular Power UP has been enabled
Public isCircularWarpMine As Boolean 'True if Circular Warp Mine has been enabled

Public SlowMotionValue As Double 'Slow Motion Value for Progress Bar
Public DistortionValue As Double 'Shield value for Progress Bar
Public EMPValue As Double 'EMP Value for Progress Bar

Public UserBitYLastPos As Double 'Variable to hold the Last known Co-Ordinate of UserBitY
Public CurrentEnemySpeed As Double 'Variable to Hold Current Enemy Speed
Public NewEnemySpeed As Double 'Variable to hold New Enemy Speed
Public ZomBitsMaxSpawnValue As Integer 'Variable to hold Maximum Spawn Value of the ZomBit
Public ZomBitsMinSpawnValue As Integer 'Variable to hold minimum Spawn Value

Public PlayerName As String 'Stores User's Name

Public UserBitBGColour As Double 'User Bit Colour
Public EnemyBitBGColour As Double 'Enemy Bit Colour
Public PowerUpBGColour As Double 'Power Up Bit Colour
Public WarpMineBGColour As Double 'Warp Mine Bit Colour
Public DistortionBGColour As Double 'Distortion Bit Colour

Public ZomBitsProgression As Integer 'Create Integer to track ZomBits Point Progress
Public EndlessSurvivalProgression As Integer 'Create Integer to Keep Stages of Endless Survival
Public MainMenuProgression As Integer 'Create Integer to Keep Main Menu Stages
Public WallNumber As Integer 'Create integer for Wall Modifier

Public ColourEditor As Integer 'Keeps track of which colour is being edited
Public TotalPoints As Integer 'Create Points
Public PointMultiplier As Integer 'Point Multiplier
