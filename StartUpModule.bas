Attribute VB_Name = "StartUpModule"

Sub StartUpMainMenu() 'Initiates Basic Things in Main Menu

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Resets
    isSlowMotion = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    MainMenuProgression = 0
    WallNumber = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    
    'Set Main Menu (X,Y) Co-ordinates
    EnemyBit1X = 0
    EnemyBit1Y = 0
    EnemyBit2X = 0
    EnemyBit2Y = 0
    EnemyBit3X = 0
    EnemyBit3Y = 0
    EnemyBit4X = 0
    EnemyBit4Y = 0
    EnemyBit5X = 0
    EnemyBit5Y = 0
    EnemyBit6X = 0
    EnemyBit6Y = 0
    WarpMineBit1X = 0
    WarpMineBit1Y = 0
    WarpMineBit2X = 0
    WarpMineBit2Y = 0
    
    'Set Main Menu Speeds
    EnemyBit1XSpeed = 0
    EnemyBit1YSpeed = 0
    EnemyBit2XSpeed = 0
    EnemyBit2YSpeed = 0
    EnemyBit3XSpeed = 0
    EnemyBit3YSpeed = 0
    EnemyBit4XSpeed = 0
    EnemyBit4YSpeed = -0
    EnemyBit5XSpeed = 0
    EnemyBit5YSpeed = 0
    EnemyBit6XSpeed = -0
    EnemyBit6YSpeed = 0
    WarpMineBit1XSpeed = 0
    WarpMineBit1YSpeed = 0
    WarpMineBit2XSpeed = -0
    WarpMineBit2YSpeed = 0
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Set points
    TotalPoints = 0
    
End Sub

Sub StartUpSurvivalEasy() 'Initiates Basic Things in Survival Mode Easy

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    SpeedIncrementation = 5
    
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Enemy Bit Positions
    EnemyBit1X = 1000
    EnemyBit1Y = 2925
    EnemyBit2X = 12000
    EnemyBit2Y = 2925
    EnemyBit3X = -300
    EnemyBit3Y = -300
    EnemyBit4X = -300
    EnemyBit4Y = -300
    EnemyBit5X = -300
    EnemyBit5Y = -300
    EnemyBit6X = -300
    EnemyBit6Y = -300
    WarpMineBit1X = -300
    WarpMineBit1Y = -300
    WarpMineBit2X = -300
    WarpMineBit2Y = -300
    
    'Enemy Bit Speeds
    EnemyBit1XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit1YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit2XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit2YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3XSpeed = 0
    EnemyBit3YSpeed = 0
    EnemyBit4XSpeed = 0
    EnemyBit4YSpeed = 0
    EnemyBit5XSpeed = 0
    EnemyBit5YSpeed = 0
    EnemyBit6XSpeed = 0
    EnemyBit6YSpeed = 0
    WarpMineBit1XSpeed = 0
    WarpMineBit1YSpeed = 0
    WarpMineBit2XSpeed = 0
    WarpMineBit2YSpeed = 0
    
    'Point Reset
    TotalPoints = 0
    PointMultiplier = 1
    
    'Resets
    isSlowMotion = False
    PowerUpPickUp = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    WallNumber = 0
    
    
End Sub

Sub StartUpSurvivalMedium() 'Initiates Basic Things in Survival Mode Medium

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    SpeedIncrementation = 5
    
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Enemy Bit Positions
    EnemyBit1X = 1000
    EnemyBit1Y = 2925
    EnemyBit2X = 12000
    EnemyBit2Y = 2925
    EnemyBit3X = 6635
    EnemyBit3Y = 500
    EnemyBit4X = -300
    EnemyBit4Y = -300
    EnemyBit5X = -300
    EnemyBit5Y = -300
    EnemyBit6X = -300
    EnemyBit6Y = -300
    WarpMineBit1X = -300
    WarpMineBit1Y = -300
    WarpMineBit2X = -300
    WarpMineBit2Y = -300
    
    'Enemy Bit Speeds
    EnemyBit1XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit1YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit2XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit2YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit4XSpeed = 0
    EnemyBit4YSpeed = 0
    EnemyBit5XSpeed = 0
    EnemyBit5YSpeed = 0
    EnemyBit6XSpeed = 0
    EnemyBit6YSpeed = 0
    WarpMineBit1XSpeed = 0
    WarpMineBit1YSpeed = 0
    WarpMineBit2XSpeed = 0
    WarpMineBit2YSpeed = 0
    
    'Point Reset
    TotalPoints = 0
    PointMultiplier = 1
    
    'Resets
    isSlowMotion = False
    PowerUpPickUp = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    WallNumber = 0
    
    
End Sub

Sub StartUpSurvivalHard() 'Initiates Basic Things in Survival Mode Hard

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    SpeedIncrementation = 10
    
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Enemy Bit Positions
    EnemyBit1X = 1000
    EnemyBit1Y = 2925
    EnemyBit2X = 12000
    EnemyBit2Y = 2925
    EnemyBit3X = 6635
    EnemyBit3Y = 500
    EnemyBit4X = 6635
    EnemyBit4Y = 5700
    EnemyBit5X = -300
    EnemyBit5Y = -300
    EnemyBit6X = -300
    EnemyBit6Y = -300
    WarpMineBit1X = -300
    WarpMineBit1Y = -300
    WarpMineBit2X = -300
    WarpMineBit2Y = -300
    
    'Enemy Bit Speeds
    EnemyBit1XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit1YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit2XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit2YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit4XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit4YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit5XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit5YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit6XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit6YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    WarpMineBit1XSpeed = 0
    WarpMineBit1YSpeed = 0
    WarpMineBit2XSpeed = 0
    WarpMineBit2YSpeed = 0
    
    'Point Reset
    TotalPoints = 0
    PointMultiplier = 1
    
    'Resets
    isSlowMotion = False
    PowerUpPickUp = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    EndlessSurvivalProgression = 0
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    WallNumber = 0
    
End Sub

Sub StartUpSurvivalEndless() 'Initiates Basic Things in Survival Mode Endless

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    SpeedIncrementation = 10
    
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Enemy Bit Positions
    EnemyBit1X = 1000
    EnemyBit1Y = 2925
    EnemyBit2X = 1000
    EnemyBit2Y = 1000
    EnemyBit3X = 1000
    EnemyBit3Y = 1000
    EnemyBit4X = 1000
    EnemyBit4Y = 1000
    EnemyBit5X = 1000
    EnemyBit5Y = 1000
    EnemyBit6X = 1000
    EnemyBit6Y = 1000
    WarpMineBit1X = 1000
    WarpMineBit1Y = 1000
    WarpMineBit2X = 1000
    WarpMineBit2Y = 1000
    
    'Enemy Bit Speeds
    EnemyBit1XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit1YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit2XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit2YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit3YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit4XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit4YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit5XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit5YSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit6XSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    EnemyBit6YSpeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    WarpMineBit1XSpeed = 0
    WarpMineBit1YSpeed = 0
    WarpMineBit2XSpeed = 0
    WarpMineBit2YSpeed = 0
    
    'Point Reset
    TotalPoints = 0
    PointMultiplier = 1
    
    'Resets
    isSlowMotion = False
    PowerUpPickUp = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    EndlessSurvivalProgression = 0
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    WallNumber = 0
    
End Sub

Sub StartUpZomBits() 'Initiates Basic Things in ZomBits

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    SpeedIncrementation = 10
    
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    ZomBitEndWallX = 600
    ZomBitEndWallY = 300
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Enemy Bit Positions
    EnemyBit1X = 13000
    EnemyBit1Y = 2925
    EnemyBit2X = 126
    EnemyBit2Y = 1000
    EnemyBit3X = 126
    EnemyBit3Y = 1000
    EnemyBit4X = 126
    EnemyBit4Y = 1000
    EnemyBit5X = 126
    EnemyBit5Y = 1000
    EnemyBit6X = 126
    EnemyBit6Y = 1000
    WarpMineBit1X = 126
    WarpMineBit1Y = 500
    WarpMineBit2X = 126
    WarpMineBit2Y = 500
    
    'Enemy Bit Speeds
    EnemyBit1XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    EnemyBit1YSpeed = 0
    EnemyBit2XSpeed = 0
    EnemyBit2YSpeed = 0
    EnemyBit3XSpeed = 0
    EnemyBit3YSpeed = 0
    EnemyBit4XSpeed = 0
    EnemyBit4YSpeed = 0
    EnemyBit5XSpeed = 0
    EnemyBit5YSpeed = 0
    EnemyBit6XSpeed = 0
    EnemyBit6YSpeed = 0
    WarpMineBit1XSpeed = 0
    WarpMineBit1YSpeed = 0
    WarpMineBit2XSpeed = 0
    WarpMineBit2YSpeed = 0
    
    'Point Reset
    TotalPoints = 0
    PointMultiplier = 1
    
    'Resets
    isSlowMotion = False
    PowerUpPickUp = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    ZomBitsMaxSpawnValue = 5700
    ZomBitsMinSpawnValue = 200
    EndlessSurvivalProgression = 0
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    WallNumber = 0
    
End Sub

Sub StartUpInstructions() 'Start Up Procedure for Instructions

    'Set User Bit (x, Y) in center of Arena
    UserBitX = UserBitXStart
    UserBitY = UserBitYStart
    
    isExploded = False
    SpeedIncrementation = 5
    
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Resets
    isSlowMotion = False
    isDistortion = False
    isEnemyDistortion = False
    isBurntOut = False
    isWall = False
    isPlayerTP = False
    isEMPOn = False
    EMPValue = 200
    SlowMotionValue = 200
    DistortionValue = 200
    WallNumber = 0
    
    'Set Speeds to Zero
    UserBitXSpeed = 0
    UserBitYSpeed = 0
    
    'Wall (X, Y) Values
    W(1) = 125
    W(2) = 125
    W(3) = 125
    W(4) = 6275
    W(5) = 125
    W(6) = 125
    W(7) = 13700
    W(8) = 125
    
    'Set Instructions (X,Y) Co-ordinates
    EnemyBit1X = 11760
    EnemyBit1Y = 3000
    EnemyBit2X = 11760
    EnemyBit2Y = 5000
    WarpMineBit1X = 11400
    WarpMineBit1Y = 1200
    
    'Set Instructions Speeds
    EnemyBit1XSpeed = 0
    EnemyBit1YSpeed = -50
    EnemyBit2XSpeed = -50
    EnemyBit2YSpeed = 0
    
    'Safety Measure Wall (X, Y) Values
    TopSafetyCounterX = 6785
    TopSafetyCounterY = 10
    LeftSafetyCounterX = 10
    LeftSafetyCounterY = 3075
    RightSafetyCounterX = 13750
    RightSafetyCounterY = 3075
    BottomSafetyCounterX = 6785
    BottomSafetyCounterY = 6350
    
    'Set points
    TotalPoints = 0

End Sub
