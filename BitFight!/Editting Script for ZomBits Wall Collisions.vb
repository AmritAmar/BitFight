Sub checkEnemyZomBitsWallCollision()

    'WALL COLLISONS
    If (ZomBitEndWallX - EnemyBit1X) > 0 Then
        
        EnemyBit1X = 13000
        EnemyBit1Y = CInt(Int((5700 * Rnd()) + 200))
	   
        EnemyBit1XSpeed = CInt(Int((-75 * Rnd()) + -50))
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If (ZomBitEndWallX - EnemyBit2X) > 0 Then
        
        EnemyBit2X = 13000
        EnemyBit2Y = CInt(Int((5700 * Rnd()) + 200))
	   
        EnemyBit2XSpeed = CInt(Int((-75 * Rnd()) + -50))
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
	If (ZomBitEndWallX - EnemyBit3X) > 0 Then
        
        EnemyBit3X = 13000
        EnemyBit3Y = CInt(Int((5700 * Rnd()) + 200))
	   
        EnemyBit3XSpeed = CInt(Int((-75 * Rnd()) + -50))
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If (ZomBitEndWallX - EnemyBit4X) > 0 Then
        
        EnemyBit4X = 13000
        EnemyBit4Y = CInt(Int((5700 * Rnd()) + 200))
	   
        EnemyBit4XSpeed = CInt(Int((-75 * Rnd()) + -50))
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
	If (ZomBitEndWallX - EnemyBit5X) > 0 Then
        
        EnemyBit5X = 13000
        EnemyBit5Y = CInt(Int((5700 * Rnd()) + 200))
	   
        EnemyBit5XSpeed = CInt(Int((-75 * Rnd()) + -50))
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If (ZomBitEndWallX - EnemyBit6X) > 0 Then
        
        EnemyBit6X = 13000
        EnemyBit6Y = CInt(Int((5700 * Rnd()) + 200))
	   
        EnemyBit6XSpeed = CInt(Int((-75 * Rnd()) + -50))
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
End Sub