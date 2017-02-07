Attribute VB_Name = "GamePlayModule"
Sub ZomBitsTargetUser(UserBitY As Integer, EnemyBitY As Integer, EnemyBitYSpeed As Integer, EnemyBitXSpeed As Integer)

    If EnemyBitY < UserBitY Then
    
        EnemyBitXSpeed = 0
        EnemyBitYSpeed = 100
    
    ElseIf EnemyBitY > UserBitY Then
    
        EnemyBitXSpeed = 0
        EnemyBitYSpeed = -100
        
    End If

End Sub
Sub checkZomBitsAiming()

    If ZomBitsProgression >= 0 Then

        If (Abs(UserBitY - EnemyBit1Y) < 250) And EnemyBit1XSpeed = 0 Then
        
            EnemyBit1YSpeed = 0
            EnemyBit1XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
            
        End If
        
        If ZomBitsProgression >= 1 Then
        
            If (Abs(UserBitY - EnemyBit2Y) < 250) And EnemyBit2XSpeed = 0 Then
        
                EnemyBit2YSpeed = 0
                EnemyBit2XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
            
            End If
            
            If ZomBitsProgression >= 2 Then
            
                If (Abs(UserBitY - EnemyBit3Y) < 250) And EnemyBit3XSpeed = 0 Then
        
                    EnemyBit3YSpeed = 0
                    EnemyBit3XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
            
                End If
                
                If ZomBitsProgression >= 4 Then
            
                    If (Abs(UserBitY - EnemyBit4Y) < 250) And EnemyBit4XSpeed = 0 Then
            
                        EnemyBit4YSpeed = 0
                        EnemyBit4XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
                
                    End If
                
                    If ZomBitsProgression >= 6 Then
                
                        If (Abs(UserBitY - EnemyBit5Y) < 250) And EnemyBit5XSpeed = 0 Then
                
                            EnemyBit5YSpeed = 0
                            EnemyBit5XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
                    
                        End If
                        
                        If ZomBitsProgression >= 7 Then
                    
                            If (Abs(UserBitY - EnemyBit6Y) < 250) And EnemyBit6XSpeed = 0 Then
                    
                                EnemyBit6YSpeed = 0
                                EnemyBit6XSpeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
                        
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End If
        
End Sub
Sub EnemyUserTarget(UserBitXPos As Integer, UserBitYPos As Integer, EnemyBitXPos As Integer, EnemyBitYPos As Integer, EnemyBitXSpeed As Integer, EnemyBitYSpeed As Integer)
'This Sub-Routine retraces the paths of Enemy Bits, and changes it towards User Bit

    If UserBitXPos < EnemyBitXPos And UserBitYPos > EnemyBitYPos Then '1st Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2)) 'Find Current Speed
        NewEnemySpeed = Sqr((Abs(EnemyBitXPos - UserBitXPos) ^ 2) + (Abs(UserBitYPos - EnemyBitYPos) ^ 2)) 'Find the new Speed
    
        'Assign New Speed in Direction of User Bit
        EnemyBitXSpeed = -CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

    If UserBitXPos < EnemyBitXPos And UserBitYPos < EnemyBitYPos Then '2nd Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2))
        NewEnemySpeed = Sqr((Abs(EnemyBitXPos - UserBitXPos) ^ 2) + (Abs(EnemyBitYPos - UserBitYPos) ^ 2))

        EnemyBitXSpeed = -CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = -CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

    If UserBitXPos > EnemyBitXPos And UserBitYPos < EnemyBitYPos Then '3rd Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2))
        NewEnemySpeed = Sqr((Abs(UserBitXPos - EnemyBitXPos) ^ 2) + (Abs(EnemyBitYPos - UserBitYPos) ^ 2))

        EnemyBitXSpeed = CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = -CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

    If UserBitXPos > EnemyBitXPos And UserBitYPos > EnemyBitYPos Then '4th Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2))
        NewEnemySpeed = Sqr((Abs(UserBitXPos - EnemyBitXPos) ^ 2) + (Abs(UserBitYPos - EnemyBitYPos) ^ 2))
    
        EnemyBitXSpeed = CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

End Sub

Sub EnemyUserLoseTarget(UserBitXPos As Integer, UserBitYPos As Integer, EnemyBitXPos As Integer, EnemyBitYPos As Integer, EnemyBitXSpeed As Integer, EnemyBitYSpeed As Integer)
'This Sub-Routine retraces the paths of Enemy Bits, and changes it from against User Bit

    If UserBitXPos < EnemyBitXPos And UserBitYPos > EnemyBitYPos Then '1st Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2)) 'Find Current Speed
        NewEnemySpeed = Sqr((Abs(EnemyBitXPos - UserBitXPos) ^ 2) + (Abs(UserBitYPos - EnemyBitYPos) ^ 2)) 'Find the new Speed
    
        'Assign New Speed in opposite Direction of User Bit
        EnemyBitXSpeed = CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = -CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

    If UserBitXPos < EnemyBitXPos And UserBitYPos < EnemyBitYPos Then '2nd Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2))
        NewEnemySpeed = Sqr((Abs(EnemyBitXPos - UserBitXPos) ^ 2) + (Abs(EnemyBitYPos - UserBitYPos) ^ 2))

        EnemyBitXSpeed = CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

    If UserBitXPos > EnemyBitXPos And UserBitYPos < EnemyBitYPos Then '3rd Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2))
        NewEnemySpeed = Sqr((Abs(UserBitXPos - EnemyBitXPos) ^ 2) + (Abs(EnemyBitYPos - UserBitYPos) ^ 2))

        EnemyBitXSpeed = -CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

    If UserBitXPos > EnemyBitXPos And UserBitYPos > EnemyBitYPos Then '4th Quadrant
        
        CurrentEnemySpeed = Sqr((EnemyBitXSpeed ^ 2) + (EnemyBitYSpeed ^ 2))
        NewEnemySpeed = Sqr((Abs(UserBitXPos - EnemyBitXPos) ^ 2) + (Abs(UserBitYPos - EnemyBitYPos) ^ 2))
    
        EnemyBitXSpeed = -CurrentEnemySpeed * (Abs(EnemyBitXPos - UserBitXPos) / NewEnemySpeed)
        EnemyBitYSpeed = -CurrentEnemySpeed * (Abs(UserBitYPos - EnemyBitYPos) / NewEnemySpeed)

    End If

End Sub

Sub CheckUserSpeeds() 'Checks User Speeds

    If UserBitXSpeed > 180 Or UserBitXSpeed < -180 Or UserBitYSpeed > 180 Or UserBitYSpeed < -180 Then 'If they are Above a Limit then
        
        isBurntOut = True 'You die!
        
    End If 'End Choices

End Sub

Sub WallModifier(WallNum As Integer) 'Wall Modifier

    If WallNum = 1 Then 'If Top Wall
    
        W(2) = 2000
    
    ElseIf WallNum = 2 Then 'If Left Wall
    
        W(5) = 3500
    
    ElseIf WallNum = 3 Then 'If Right Wall
    
        W(7) = 10000
    
    ElseIf WallNum = 4 Then 'If Bottom Wall
    
        W(4) = 4500
    
    End If 'End Selection
    
End Sub
Sub ZomBitsWallModifier(WallNum As Integer) 'Wall Modifier

    If WallNum = 1 Then 'If Top Wall
    
        W(2) = 2000
    
    ElseIf WallNum = 2 Then 'If Bottom Wall
    
        W(4) = 4500
    
    End If 'End Selection
    
End Sub

Sub WallBitPosition(WallNum As Integer, EnemyBits As Integer, MineBits As Integer)
'Checks Wall Number and if objects are outside the bounds, then it repositions them

    If WallNum = 1 Then 'If TOP Wall
    
        If EnemyBits = 2 Then 'If 2 Enemies Only
            
            'Checks the First 2 Enemies if They Are Out
            If EnemyBit1Y < (W(2) + 80) Then
            
                EnemyBit1Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit2Y < W(2) Then
            
                EnemyBit2Y = W(2) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY < W(2) Then
            
                UserBitY = W(2) + 100 'Repositions
                
            End If
        
        End If 'End 2 ENEMY Check
        
        If EnemyBits = 3 Then 'If 3 Enemies Only
            
            'Checks the First 3 Enemies if They Are Out
            If EnemyBit1Y < (W(2) + 80) Then
            
                EnemyBit1Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit2Y < W(2) Then
            
                EnemyBit2Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit3Y < W(2) Then
            
                EnemyBit3Y = W(2) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY < W(2) Then
            
                UserBitY = W(2) + 100 'Repositions
                
            End If
        
        End If 'End 3 ENEMY Check
        
        If EnemyBits = 4 Then 'If 4 Enemies Only
            
            'Checks the First 4 Enemies if They Are Out
            If EnemyBit1Y < (W(2) + 80) Then
            
                EnemyBit1Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit2Y < W(2) Then
            
                EnemyBit2Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit3Y < W(2) Then
            
                EnemyBit3Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit4Y < W(2) Then
            
                EnemyBit4Y = W(2) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY < W(2) Then
            
                UserBitY = W(2) + 100 'Repositions
                
            End If
        
        End If 'End 4 ENEMY Check
        
        If EnemyBits = 6 Then 'If 6 Enemies Only
            
            'Checks the First 6 Enemies if They Are Out
            If EnemyBit1Y < (W(2) + 80) Then
            
                EnemyBit1Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit2Y < W(2) Then
            
                EnemyBit2Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit3Y < W(2) Then
            
                EnemyBit3Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit4Y < W(2) Then
            
                EnemyBit4Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit5Y < W(2) Then
            
                EnemyBit5Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit6Y < W(2) Then
            
                EnemyBit6Y = W(2) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY < W(2) Then
            
                UserBitY = W(2) + 100 'Repositions
                
            End If
        
        End If 'End 6 ENEMY Check
    
    ElseIf WallNum = 2 Then 'If LEFT Wall
        
        If EnemyBits = 2 Then 'If 2 Enemies Only
            
            'Checks the First 2 Enemies if They Are Out
            If EnemyBit1X < (W(5) + 80) Then
            
                EnemyBit1X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit2X < W(5) Then
            
                EnemyBit2X = W(5) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX < W(5) Then
            
                UserBitX = W(5) + 100 'Repositions
                
            End If
        
        End If 'End 2 ENEMY Check
        
        If EnemyBits = 3 Then 'If 3 Enemies Only
            
            'Checks the First 3 Enemies if They Are Out
            If EnemyBit1X < (W(5) + 80) Then
            
                EnemyBit1X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit2X < W(5) Then
            
                EnemyBit2X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit3X < W(5) Then
            
                EnemyBit3X = W(5) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX < W(5) Then
            
                UserBitX = W(5) + 100 'Repositions
                
            End If
        
        End If 'End 3 ENEMY Check
    
        If EnemyBits = 4 Then  'If 4 Enemies Only
            
            'Checks the First 4 Enemies if They Are Out
            If EnemyBit1X < (W(5) + 80) Then
            
                EnemyBit1X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit2X < W(5) Then
            
                EnemyBit2X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit3X < W(5) Then
            
                EnemyBit3X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit4X < W(5) Then
            
                EnemyBit4X = W(5) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX < W(5) Then
            
                UserBitX = W(5) + 100 'Repositions
                
            End If
        
        End If 'End 4 ENEMY Check
        
        If EnemyBits = 6 Then  'If 6 Enemies Only
            
            'Checks the First 6 Enemies if They Are Out
            If EnemyBit1X < (W(5) + 80) Then
            
                EnemyBit1X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit2X < W(5) Then
            
                EnemyBit2X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit3X < W(5) Then
            
                EnemyBit3X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit4X < W(5) Then
            
                EnemyBit4X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit5X < W(5) Then
            
                EnemyBit5X = W(5) + 100 'Repositions
                
            End If
            
            If EnemyBit6X < W(5) Then
            
                EnemyBit6X = W(5) + 100 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX < W(5) Then
            
                UserBitX = W(5) + 100 'Repositions
                
            End If
        
        End If 'End 6 ENEMY Check
        
    ElseIf WallNum = 3 Then 'if RIGHT Wall
        
        If EnemyBits = 2 Then 'If 2 Enemies Only
            
            'Checks the First 2 Enemies if They Are Out
            If EnemyBit1X > (W(7) - 80) Then
            
                EnemyBit1X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit2X > (W(7) - 340) Then
            
                EnemyBit2X = W(7) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX > (W(7) - 340) Then
            
                UserBitX = W(7) - 350 'Repositions
                
            End If
        
        End If 'End 2 ENEMY Check
        
        If EnemyBits = 3 Then 'If 3 Enemies Only
            
            'Checks the First 3 Enemies if They Are Out
            If EnemyBit1X > (W(7) - 80) Then
            
                EnemyBit1X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit2X > (W(7) - 340) Then
            
                EnemyBit2X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit3X > (W(7) - 340) Then
            
                EnemyBit3X = W(7) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX > (W(7) - 340) Then
            
                UserBitX = W(7) - 350 'Repositions
                
            End If
        
        End If 'End 3 ENEMY Check
        
        If EnemyBits = 4 Then 'If 4 Enemies Only
            
            'Checks the First 4 Enemies if They Are Out
            If EnemyBit1X > (W(7) - 80) Then
            
                EnemyBit1X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit2X > (W(7) - 340) Then
            
                EnemyBit2X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit3X > (W(7) - 340) Then
            
                EnemyBit3X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit4X > (W(7) - 340) Then
            
                EnemyBit4X = W(7) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX > W(7) Then
            
                UserBitX = W(7) - 350 'Repositions
                
            End If
        
        End If 'End 4 ENEMY Check
        
        If EnemyBits = 6 Then 'If 6 Enemies Only
            
            'Checks the First 6 Enemies if They Are Out
            If EnemyBit1X > (W(7) - 80) Then
            
                EnemyBit1X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit2X > (W(7) - 340) Then
            
                EnemyBit2X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit3X > (W(7) - 340) Then
            
                EnemyBit3X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit4X > (W(7) - 340) Then
            
                EnemyBit4X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit5X > (W(7) - 340) Then
            
                EnemyBit5X = W(7) - 350 'Repositions
                
            End If
            
            If EnemyBit6X > (W(7) - 340) Then
            
                EnemyBit6X = W(7) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitX > (W(7) - 340) Then
            
                UserBitX = W(7) - 350 'Repositions
                
            End If
        
        End If 'End 6 ENEMY Check
    
    ElseIf WallNum = 4 Then 'if BOTTOM Wall
    
        If EnemyBits = 2 Then 'If 2 Enemies Only
            
            'Checks the First 2 Enemies if They Are Out
            If EnemyBit1Y > (W(4) - 80) Then
            
                EnemyBit1Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit2Y > (W(4) - 340) Then
            
                EnemyBit2Y = W(4) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY > (W(4) - 340) Then
            
                UserBitY = W(4) - 350 'Repositions
                
            End If
        
        End If 'End 2 ENEMY Check
        
        If EnemyBits = 3 Then 'If 3 Enemies Only
            
            'Checks the First 3 Enemies if They Are Out
            If EnemyBit1Y > (W(4) - 340) Then
            
                EnemyBit1Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit2Y > (W(4) - 340) Then
            
                EnemyBit2Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit3Y > (W(4) - 340) Then
            
                EnemyBit3Y = W(4) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY > (W(4) - 340) Then
            
                UserBitY = W(4) - 350 'Repositions
                
            End If
        
        End If 'End 3 ENEMY Check
    
        If EnemyBits = 4 Then 'If 4 Enemies Only
            
            'Checks the First 4 Enemies if They Are Out
            If EnemyBit1Y > (W(4) - 340) Then
            
                EnemyBit1Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit2Y > (W(4) - 340) Then
            
                EnemyBit2Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit3Y > (W(4) - 340) Then
            
                EnemyBit3Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit4Y > (W(4) - 340) Then
            
                EnemyBit4Y = W(4) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY > (W(4) - 340) Then
            
                UserBitY = W(4) - 350 'Repositions
                
            End If
        
        End If 'End 4 ENEMY Check
    
        If EnemyBits = 6 Then 'If 6 Enemies Only
            
            'Checks the First 6 Enemies if They Are Out
            If EnemyBit1Y > (W(4) - 340) Then
            
                EnemyBit1Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit2Y > (W(4) - 340) Then
            
                EnemyBit2Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit3Y > (W(4) - 340) Then
            
                EnemyBit3Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit4Y > (W(4) - 340) Then
            
                EnemyBit4Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit5Y > (W(4) - 340) Then
            
                EnemyBit5Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit6Y > (W(4) - 340) Then
            
                EnemyBit6Y = W(4) - 350 'Repositions
                
            End If
            
            'Checks the User if it is Out
            If UserBitY > (W(4) - 340) Then
            
                UserBitY = W(4) - 350 'Repositions
                
            End If
        
        End If 'End 6 ENEMY Check
    
    End If

End Sub
Sub ZomBitWallBitPosition(WallNum As Integer)
'Checks Wall Number and if objects are outside the bounds, then it repositions them

    If WallNum = 1 Then 'If TOP Wall
    
        ZomBitsMinSpawnValue = 2200
        ZomBitsMaxSpawnValue = 5700
            
            If EnemyBit1Y < (W(2) + 80) Then
            
                EnemyBit1Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit2Y < W(2) Then
            
                EnemyBit2Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit3Y < W(2) Then
            
                EnemyBit3Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit4Y < W(2) Then
            
                EnemyBit4Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit5Y < W(2) Then
            
                EnemyBit5Y = W(2) + 100 'Repositions
                
            End If
            
            If EnemyBit6Y < W(2) Then
            
                EnemyBit6Y = W(2) + 100 'Repositions
                
            End If
            
        'Checks the User if it is Out
        If UserBitY < W(2) Then
        
            UserBitY = W(2) + 100 'Repositions
            
        End If
        
    ElseIf WallNum = 2 Then 'If BOTTOM Wall
    
        ZomBitsMaxSpawnValue = 4100
        
            If EnemyBit1Y > (W(4) - 340) Then
            
                EnemyBit1Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit2Y > (W(4) - 340) Then
            
                EnemyBit2Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit3Y > (W(4) - 340) Then
            
                EnemyBit3Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit4Y > (W(4) - 340) Then
            
                EnemyBit4Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit5Y > (W(4) - 340) Then
            
                EnemyBit5Y = W(4) - 350 'Repositions
                
            End If
            
            If EnemyBit6Y > (W(4) - 340) Then
            
                EnemyBit6Y = W(4) - 350 'Repositions
                
            End If
        
        'Checks the User if it is Out
        If UserBitY > (W(4) - 340) Then
        
            UserBitY = W(4) - 350 'Repositions
            
        End If
        
    End If
    
End Sub

Sub checkWarpBitSpawnProximity() 'Checks if The Warp Mine has spawned to close to the player

    If Abs(WarpMineBit1X - UserBitX) < 600 And Abs(WarpMineBit1Y - UserBitY) < 600 Then
    
        isWarpMineProximal = True
    
    End If
    
    If Abs(WarpMineBit2X - UserBitX) < 600 And Abs(WarpMineBit2Y - UserBitY) < 600 Then
    
        isWarpMineProximal = True
    
    End If

End Sub

Sub EnemyIncreaseSpeed(EnemyBitXSpeed, EnemyBitYSpeed) 'Increases the Enemy Speed

    If EnemyBitXSpeed < 85 Then
                
        EnemyBitXSpeed = EnemyBitXSpeed + (EnemyBitXSpeed / 10) 'Enemy Bit X Speed +10%
    
    End If
    
    If EnemyBitYSpeed < 85 Then
    
        EnemyBitYSpeed = EnemyBitYSpeed + (EnemyBitYSpeed / 10) 'Enemy Bit 1 Y Speed +10%
    
    End If

End Sub

Sub EnemyDecreaseSpeed(EnemyBitXSpeed, EnemyBitYSpeed) 'Increases the Enemy Speed

    EnemyBitXSpeed = EnemyBitXSpeed - (EnemyBitXSpeed / 10) 'Enemy Bit X Speed +10%
    EnemyBitYSpeed = EnemyBitYSpeed - (EnemyBitYSpeed / 10) 'Enemy Bit Y Speed +10%

End Sub
