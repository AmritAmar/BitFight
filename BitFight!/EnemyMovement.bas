Attribute VB_Name = "EnemyMovement"
Sub EnemyBitMovement(EnemyBitX As Integer, EnemyBitY As Integer, EnemyBitXSpeed As Integer, EnemyBitYSpeed As Integer) 'Set Movement and Mapping
    
    If isExploded = False Then
    
        'Set Enemy Bits to Move
        EnemyBitX = EnemyBitX + EnemyBitXSpeed
        EnemyBitY = EnemyBitY + EnemyBitYSpeed

    End If
    
End Sub

Sub WarpMineBitPosition(WarpMineBitX As Integer, WarpMineBitY As Integer)

    If isExploded = False Then
    
        'Assigns a Random (X, Y) position to Warp Mines
        WarpMineBitX = CInt(Int((12600 - 300 + 1) * Rnd() + 300))
        WarpMineBitY = CInt(Int((5500 - 300 + 1) * Rnd() + 300))
    
    End If

End Sub
