Attribute VB_Name = "CollisonModule"
Sub checkUserZomBitWallCollision()

    If isDistortion = False Then
    
        If ZomBitEndWallX > UserBitX Then
            
            isEnemyHit = True
        
        End If
        
    End If

End Sub

Sub checkEnemyZomBitsWallCollision(Progression As Integer)

    'WALL COLLISONs
    If Progression >= 0 Then
        
        If (ZomBitEndWallX - EnemyBit1X) > 0 Then
            
            EnemyBit1X = 13000
            EnemyBit1Y = CInt(Int((ZomBitsMaxSpawnValue * Rnd()) + ZomBitsMinSpawnValue))
           
            EnemyBit1XSpeed = CInt(Int((-75 * Rnd()) + -50))
            
            TotalPoints = TotalPoints + ((1) * PointMultiplier)
        
        End If
        
        If Abs(W(1) - EnemyBit1X) < 13572 And Abs(W(2) - EnemyBit1Y) < 49 Then
                
            EnemyBit1YSpeed = (EnemyBit1YSpeed * -1)
            EnemyBit1Y = W(1) + 50
            TotalPoints = TotalPoints + ((1) * PointMultiplier)
            
        End If
            
        If Abs(W(3) - EnemyBit1X) < 13572 And Abs(W(4) - EnemyBit1Y) < 349 Then
                
            EnemyBit1YSpeed = (EnemyBit1YSpeed * -1)
            EnemyBit1Y = W(4) - 350
            TotalPoints = TotalPoints + ((1) * PointMultiplier)
            
        End If
            
        If Abs(W(5) - EnemyBit1X) < 49 And Abs(W(6) - EnemyBit1Y) < 6156 Then
                
            EnemyBit1XSpeed = (EnemyBit1XSpeed * -1)
            EnemyBit1X = W(5) + 50
            TotalPoints = TotalPoints + ((1) * PointMultiplier)
            
        End If
            
        If Abs(W(7) - EnemyBit1X) < 349 And Abs(W(8) - EnemyBit1Y) < 6156 Then
                
            EnemyBit1XSpeed = (EnemyBit1XSpeed * -1)
            EnemyBit1X = W(7) - 350
            TotalPoints = TotalPoints + ((1) * PointMultiplier)
            
        End If
        
        If (TopSafetyCounterY - EnemyBit1Y) > -(50) Then
    
        EnemyBit1YSpeed = EnemyBit1YSpeed * -1
        EnemyBit1Y = EnemyBit1Y + 150
        
        End If
        
        If (BottomSafetyCounterY - EnemyBit1Y) < 25 Then
        
            EnemyBit1YSpeed = EnemyBit1YSpeed * -1
            EnemyBit1Y = EnemyBit1Y - 150
        
        End If
        
        If (LeftSafetyCounterX - EnemyBit1X) > -50 Then
        
            EnemyBit1XSpeed = EnemyBit1XSpeed * -1
            EnemyBit1X = EnemyBit1X + 150
        
        End If
        
        If (RightSafetyCounterX - EnemyBit1X) < 25 Then
        
            EnemyBit1XSpeed = EnemyBit1XSpeed * -1
            EnemyBit1X = EnemyBit1X - 150
        
        End If
        
        If Progression >= 1 Then
            
            If (ZomBitEndWallX - EnemyBit2X) > 0 Then
                
                EnemyBit2X = 13000
                EnemyBit2Y = CInt(Int((ZomBitsMaxSpawnValue * Rnd()) + ZomBitsMinSpawnValue))
               
                EnemyBit2XSpeed = CInt(Int((-75 * Rnd()) + -50))
                
                TotalPoints = TotalPoints + ((1) * PointMultiplier)
            
            End If
                        
            If Abs(W(1) - EnemyBit2X) < 13572 And Abs(W(2) - EnemyBit2Y) < 49 Then
                    
                EnemyBit2YSpeed = (EnemyBit2YSpeed * -1)
                EnemyBit2Y = W(1) + 50
                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
            End If
                
            If Abs(W(3) - EnemyBit2X) < 13572 And Abs(W(4) - EnemyBit2Y) < 349 Then
                    
                EnemyBit2YSpeed = (EnemyBit2YSpeed * -1)
                EnemyBit2Y = W(4) - 350
                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
            End If
                
            If Abs(W(5) - EnemyBit2X) < 49 And Abs(W(6) - EnemyBit2Y) < 6156 Then
                    
                EnemyBit2XSpeed = (EnemyBit2XSpeed * -1)
                EnemyBit2X = W(5) + 50
                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
            End If
                
            If Abs(W(7) - EnemyBit2X) < 349 And Abs(W(8) - EnemyBit2Y) < 6156 Then
                    
                EnemyBit2XSpeed = (EnemyBit2XSpeed * -1)
                EnemyBit2X = W(7) - 350
                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
            End If
            
            If (TopSafetyCounterY - EnemyBit2Y) > -(50) Then

                EnemyBit2YSpeed = EnemyBit2YSpeed * -1
                EnemyBit2Y = EnemyBit2Y + 150
                
            End If
            
            If (BottomSafetyCounterY - EnemyBit2Y) < 25 Then
            
                EnemyBit2YSpeed = EnemyBit2YSpeed * -1
                EnemyBit2Y = EnemyBit2Y - 150
            
            End If
            
            If (LeftSafetyCounterX - EnemyBit2X) > -50 Then
            
                EnemyBit2XSpeed = EnemyBit2XSpeed * -1
                EnemyBit2X = EnemyBit2X + 150
            
            End If
            
            If (RightSafetyCounterX - EnemyBit2X) < 25 Then
            
                EnemyBit2XSpeed = EnemyBit2XSpeed * -1
                EnemyBit2X = EnemyBit2X - 150
            
            End If
                        
            If Progression >= 2 Then
            
                If (ZomBitEndWallX - EnemyBit3X) > 0 Then
                    
                    EnemyBit3X = 13000
                    EnemyBit3Y = CInt(Int((ZomBitsMaxSpawnValue * Rnd()) + ZomBitsMinSpawnValue))
                   
                    EnemyBit3XSpeed = CInt(Int((-75 * Rnd()) + -50))
                    
                    TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
                End If
                
                If Abs(W(1) - EnemyBit3X) < 13572 And Abs(W(2) - EnemyBit3Y) < 49 Then
                    
                    EnemyBit3YSpeed = (EnemyBit3YSpeed * -1)
                    EnemyBit3Y = W(1) + 50
                    TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
                End If
                
                If Abs(W(3) - EnemyBit3X) < 13572 And Abs(W(4) - EnemyBit3Y) < 349 Then
                    
                    EnemyBit3YSpeed = (EnemyBit3YSpeed * -1)
                    EnemyBit3Y = W(4) - 350
                    TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
                End If
                
                If Abs(W(5) - EnemyBit3X) < 49 And Abs(W(6) - EnemyBit3Y) < 6156 Then
                    
                    EnemyBit3XSpeed = (EnemyBit3XSpeed * -1)
                    EnemyBit3X = W(5) + 50
                    TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
                End If
                
                If Abs(W(7) - EnemyBit3X) < 349 And Abs(W(8) - EnemyBit3Y) < 6156 Then
                    
                   EnemyBit3XSpeed = (EnemyBit3XSpeed * -1)
                   EnemyBit3X = W(7) - 350
                   TotalPoints = TotalPoints + ((1) * PointMultiplier)
                
                End If
                
                If (TopSafetyCounterY - EnemyBit3Y) > -(50) Then
                
                    EnemyBit3YSpeed = EnemyBit3YSpeed * -1
                    EnemyBit3Y = EnemyBit3Y + 150
                    
                End If
                
                If (BottomSafetyCounterY - EnemyBit3Y) < 25 Then
                
                    EnemyBit3YSpeed = EnemyBit3YSpeed * -1
                    EnemyBit3Y = EnemyBit3Y - 150
                
                End If
                
                If (LeftSafetyCounterX - EnemyBit3X) > -50 Then
                
                    EnemyBit3XSpeed = EnemyBit3XSpeed * -1
                    EnemyBit3X = EnemyBit3X + 150
                
                End If
                
                If (RightSafetyCounterX - EnemyBit3X) < 25 Then
                
                    EnemyBit3XSpeed = EnemyBit3XSpeed * -1
                    EnemyBit3X = EnemyBit3X - 150
                
                End If
                
                If Progression >= 4 Then
                
                    If (ZomBitEndWallX - EnemyBit4X) > 0 Then
                        
                        EnemyBit4X = 13000
                        EnemyBit4Y = CInt(Int((ZomBitsMaxSpawnValue * Rnd()) + ZomBitsMinSpawnValue))
                       
                        EnemyBit4XSpeed = CInt(Int((-75 * Rnd()) + -50))
                        
                        TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                    End If
                    
                    If Abs(W(1) - EnemyBit4X) < 13572 And Abs(W(2) - EnemyBit4Y) < 49 Then
        
                            EnemyBit4YSpeed = (EnemyBit4YSpeed * -1)
                            EnemyBit4Y = W(1) + 50
                            TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                    End If
                        
                    If Abs(W(3) - EnemyBit4X) < 13572 And Abs(W(4) - EnemyBit4Y) < 349 Then
                            
                        EnemyBit4YSpeed = (EnemyBit4YSpeed * -1)
                        EnemyBit4Y = W(4) - 350
                        TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                    End If
                        
                    If Abs(W(5) - EnemyBit4X) < 49 And Abs(W(6) - EnemyBit4Y) < 6156 Then
                            
                        EnemyBit4XSpeed = (EnemyBit4XSpeed * -1)
                        EnemyBit4X = W(5) + 50
                        TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                    End If
                        
                    If Abs(W(7) - EnemyBit4X) < 349 And Abs(W(8) - EnemyBit4Y) < 6156 Then
                            
                        EnemyBit4XSpeed = (EnemyBit4XSpeed * -1)
                        EnemyBit4X = W(7) - 350
                        TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                    End If
                    
                    If (TopSafetyCounterY - EnemyBit4Y) > -(50) Then
    
                        EnemyBit4YSpeed = EnemyBit4YSpeed * -1
                        EnemyBit4Y = EnemyBit4Y + 150
                        
                    End If
                    
                    If (BottomSafetyCounterY - EnemyBit4Y) < 25 Then
                    
                        EnemyBit4YSpeed = EnemyBit4YSpeed * -1
                        EnemyBit4Y = EnemyBit4Y - 150
                    
                    End If
                    
                    If (LeftSafetyCounterX - EnemyBit4X) > -50 Then
                    
                        EnemyBit4XSpeed = EnemyBit4XSpeed * -1
                        EnemyBit4X = EnemyBit4X + 150
                    
                    End If
                    
                    If (RightSafetyCounterX - EnemyBit4X) < 25 Then
                    
                        EnemyBit4XSpeed = EnemyBit4XSpeed * -1
                        EnemyBit4X = EnemyBit4X - 150
                    
                    End If
                    
                    If Progression >= 6 Then
                    
                        If (ZomBitEndWallX - EnemyBit5X) > 0 Then
                            
                            EnemyBit5X = 13000
                            EnemyBit5Y = CInt(Int((ZomBitsMaxSpawnValue * Rnd()) + ZomBitsMinSpawnValue))
                           
                            EnemyBit5XSpeed = CInt(Int((-75 * Rnd()) + -50))
                            EnemyBit5YSpeed = CInt(Int(75 * Rnd()) + -75)
                            
                            TotalPoints = TotalPoints + ((1) * PointMultiplier)
                            
                        End If
                        
                        If Abs(W(1) - EnemyBit5X) < 13572 And Abs(W(2) - EnemyBit5Y) < 49 Then
        
                            EnemyBit5YSpeed = (EnemyBit5YSpeed * -1)
                            EnemyBit5Y = W(1) + 50
                            TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                        End If
                        
                        If Abs(W(3) - EnemyBit5X) < 13572 And Abs(W(4) - EnemyBit5Y) < 349 Then
                            
                            EnemyBit5YSpeed = (EnemyBit5YSpeed * -1)
                            EnemyBit5Y = W(4) - 350
                            TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                        End If
                        
                        If Abs(W(5) - EnemyBit5X) < 49 And Abs(W(6) - EnemyBit5Y) < 6156 Then
                            
                            EnemyBit5XSpeed = (EnemyBit5XSpeed * -1)
                            EnemyBit5X = W(5) + 50
                            TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                        End If
                        
                        If Abs(W(7) - EnemyBit5X) < 349 And Abs(W(8) - EnemyBit5Y) < 6156 Then
                            
                           EnemyBit5XSpeed = (EnemyBit5XSpeed * -1)
                           EnemyBit5X = W(7) - 350
                           TotalPoints = TotalPoints + ((1) * PointMultiplier)
                        
                        End If
                        
                        If Progression >= 7 Then
                        
                            If (ZomBitEndWallX - EnemyBit6X) > 0 Then
                                
                                EnemyBit6X = 13000
                                EnemyBit6Y = CInt(Int((ZomBitsMaxSpawnValue * Rnd()) + ZomBitsMinSpawnValue))
                               
                                EnemyBit6XSpeed = CInt(Int((-75 * Rnd()) + -50))
                                EnemyBit6YSpeed = CInt(Int(75 * Rnd()) + -75)
                                
                                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                            
                            End If
                            
                            If Abs(W(1) - EnemyBit6X) < 13572 And Abs(W(2) - EnemyBit6Y) < 49 Then
        
                                EnemyBit6YSpeed = (EnemyBit6YSpeed * -1)
                                EnemyBit6Y = W(1) + 50
                                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                            
                            End If
                            
                            If Abs(W(3) - EnemyBit6X) < 13572 And Abs(W(4) - EnemyBit6Y) < 349 Then
                                
                                EnemyBit6YSpeed = (EnemyBit6YSpeed * -1)
                                EnemyBit6Y = W(4) - 350
                                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                            
                            End If
                            
                            If Abs(W(5) - EnemyBit6X) < 49 And Abs(W(6) - EnemyBit6Y) < 6156 Then
                                
                                EnemyBit6XSpeed = (EnemyBit6XSpeed * -1)
                                EnemyBit6X = W(5) + 50
                                TotalPoints = TotalPoints + ((1) * PointMultiplier)
                            
                            End If
                            
                            If Abs(W(7) - EnemyBit6X) < 349 And Abs(W(8) - EnemyBit6Y) < 6156 Then
                                
                               EnemyBit6XSpeed = (EnemyBit6XSpeed * -1)
                               EnemyBit6X = W(7) - 350
                               TotalPoints = TotalPoints + ((1) * PointMultiplier)
                            
                            End If
                            
                            If (TopSafetyCounterY - EnemyBit6Y) > -(50) Then
                            
                                EnemyBit6YSpeed = EnemyBit6YSpeed * -1
                                EnemyBit6Y = EnemyBit6Y + 150
                                
                            End If
                            
                            If (BottomSafetyCounterY - EnemyBit6Y) < 25 Then
                            
                                EnemyBit6YSpeed = EnemyBit6YSpeed * -1
                                EnemyBit6Y = EnemyBit6Y - 150
                            
                            End If
                            
                            If (LeftSafetyCounterX - EnemyBit6X) > -50 Then
                            
                                EnemyBit6XSpeed = EnemyBit6XSpeed * -1
                                EnemyBit6X = EnemyBit6X + 150
                            
                            End If
                            
                            If (RightSafetyCounterX - EnemyBit6X) < 25 Then
                            
                                EnemyBit6XSpeed = EnemyBit6XSpeed * -1
                                EnemyBit6X = EnemyBit6X - 150
                            
                            End If
                            
                        End If
                    
                    End If
                    
                End If
                
            End If
            
        End If
        
    End If
    
End Sub

Sub checkUserWallCollison() 'Checks if User has hit Wall
    
    'WALL COLLISONS
    If Abs(W(1) - UserBitX) < 13572 And Abs(W(2) - UserBitY) < 49 Then
        
        UserBitYSpeed = (UserBitYSpeed * -1)
        UserBitY = UserBitY + 50
    
    End If
    
    If Abs(W(3) - UserBitX) < 13572 And Abs(W(4) - UserBitY) < 349 Then
        
        UserBitYSpeed = (UserBitYSpeed * -1)
        UserBitY = W(4) - 350
    
    End If
    
    If Abs(W(5) - UserBitX) < 49 And Abs(W(6) - UserBitY) < 6156 Then
        
        UserBitXSpeed = (UserBitXSpeed * -1)
        UserBitX = W(5) + 50
    
    End If
    
    If Abs(W(7) - UserBitX) < 349 And Abs(W(8) - UserBitY) < 6156 Then
        
        UserBitXSpeed = (UserBitXSpeed * -1)
        UserBitX = W(7) - 350
    
    End If
    
    'SAFETY MEASURES AGAINST OVERFLOW ERRORS
    If (TopSafetyCounterY - UserBitY) > -(50) Then
    
        UserBitYSpeed = UserBitYSpeed * -1
        UserBitY = UserBitY + 150
        
    End If
    
    If (BottomSafetyCounterY - UserBitY) < 25 Then
    
        UserBitYSpeed = UserBitYSpeed * -1
        UserBitY = UserBitY - 150
    
    End If
    
    If (LeftSafetyCounterX - UserBitX) > -50 Then
    
        UserBitXSpeed = UserBitXSpeed * -1
        UserBitX = UserBitX + 150
    
    End If
    
    If (RightSafetyCounterX - UserBitX) < 25 Then
    
        UserBitXSpeed = UserBitXSpeed * -1
        UserBitX = UserBitX - 150
    
    End If

End Sub
Sub checkEnemyWallCollision()

    'WALL COLLISONS
    If Abs(W(1) - EnemyBit1X) < 13572 And Abs(W(2) - EnemyBit1Y) < 49 Then
        
        EnemyBit1YSpeed = (EnemyBit1YSpeed * -1)
        EnemyBit1Y = EnemyBit1Y + 20
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - EnemyBit1X) < 13572 And Abs(W(4) - EnemyBit1Y) < 349 Then
        
        EnemyBit1YSpeed = (EnemyBit1YSpeed * -1)
        EnemyBit1Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - EnemyBit1X) < 49 And Abs(W(6) - EnemyBit1Y) < 6156 Then
        
        EnemyBit1XSpeed = (EnemyBit1XSpeed * -1)
        EnemyBit1X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - EnemyBit1X) < 349 And Abs(W(8) - EnemyBit1Y) < 6156 Then
        
       EnemyBit1XSpeed = (EnemyBit1XSpeed * -1)
       EnemyBit1X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - EnemyBit2X) < 13572 And Abs(W(2) - EnemyBit2Y) < 49 Then
        
        EnemyBit2YSpeed = (EnemyBit2YSpeed * -1)
        EnemyBit2Y = EnemyBit2Y + 20
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - EnemyBit2X) < 13572 And Abs(W(4) - EnemyBit2Y) < 349 Then
        
        EnemyBit2YSpeed = (EnemyBit2YSpeed * -1)
        EnemyBit2Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - EnemyBit2X) < 49 And Abs(W(6) - EnemyBit2Y) < 6156 Then
        
        EnemyBit2XSpeed = (EnemyBit2XSpeed * -1)
        EnemyBit2X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - EnemyBit2X) < 349 And Abs(W(8) - EnemyBit2Y) < 6156 Then
        
       EnemyBit2XSpeed = (EnemyBit2XSpeed * -1)
       EnemyBit2X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - EnemyBit3X) < 13572 And Abs(W(2) - EnemyBit3Y) < 49 Then
        
        EnemyBit3YSpeed = (EnemyBit3YSpeed * -1)
        EnemyBit3Y = EnemyBit3Y + 20
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - EnemyBit3X) < 13572 And Abs(W(4) - EnemyBit3Y) < 349 Then
        
        EnemyBit3YSpeed = (EnemyBit3YSpeed * -1)
        EnemyBit3Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - EnemyBit3X) < 49 And Abs(W(6) - EnemyBit3Y) < 6156 Then
        
        EnemyBit3XSpeed = (EnemyBit3XSpeed * -1)
        EnemyBit3X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - EnemyBit3X) < 349 And Abs(W(8) - EnemyBit3Y) < 6156 Then
        
       EnemyBit3XSpeed = (EnemyBit3XSpeed * -1)
       EnemyBit3X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - EnemyBit4X) < 13572 And Abs(W(2) - EnemyBit4Y) < 49 Then
        
        EnemyBit4YSpeed = (EnemyBit4YSpeed * -1)
        EnemyBit4Y = EnemyBit4Y + 20
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - EnemyBit4X) < 13572 And Abs(W(4) - EnemyBit4Y) < 349 Then
        
        EnemyBit4YSpeed = (EnemyBit4YSpeed * -1)
        EnemyBit4Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - EnemyBit4X) < 49 And Abs(W(6) - EnemyBit4Y) < 6156 Then
        
        EnemyBit4XSpeed = (EnemyBit4XSpeed * -1)
        EnemyBit4X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - EnemyBit4X) < 349 And Abs(W(8) - EnemyBit4Y) < 6156 Then
        
       EnemyBit4XSpeed = (EnemyBit4XSpeed * -1)
       EnemyBit4X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - EnemyBit5X) < 13572 And Abs(W(2) - EnemyBit5Y) < 49 Then
        
        EnemyBit5YSpeed = (EnemyBit5YSpeed * -1)
        EnemyBit5Y = EnemyBit5Y + 20
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - EnemyBit5X) < 13572 And Abs(W(4) - EnemyBit5Y) < 349 Then
        
        EnemyBit5YSpeed = (EnemyBit5YSpeed * -1)
        EnemyBit5Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - EnemyBit5X) < 49 And Abs(W(6) - EnemyBit5Y) < 6156 Then
        
        EnemyBit5XSpeed = (EnemyBit5XSpeed * -1)
        EnemyBit5X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - EnemyBit5X) < 349 And Abs(W(8) - EnemyBit5Y) < 6156 Then
        
       EnemyBit5XSpeed = (EnemyBit5XSpeed * -1)
       EnemyBit5X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - EnemyBit6X) < 13572 And Abs(W(2) - EnemyBit6Y) < 49 Then
        
        EnemyBit6YSpeed = (EnemyBit6YSpeed * -1)
        EnemyBit6Y = EnemyBit6Y + 20
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - EnemyBit6X) < 13572 And Abs(W(4) - EnemyBit6Y) < 349 Then
        
        EnemyBit6YSpeed = (EnemyBit6YSpeed * -1)
        EnemyBit6Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - EnemyBit6X) < 49 And Abs(W(6) - EnemyBit6Y) < 6156 Then
        
        EnemyBit6XSpeed = (EnemyBit6XSpeed * -1)
        EnemyBit6X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - EnemyBit6X) < 349 And Abs(W(8) - EnemyBit6Y) < 6156 Then
        
       EnemyBit6XSpeed = (EnemyBit6XSpeed * -1)
       EnemyBit6X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - WarpMineBit1X) < 13572 And Abs(W(2) - WarpMineBit1Y) < 49 Then
        
        WarpMineBit1YSpeed = (WarpMineBit1YSpeed * -1)
        WarpMineBit1Y = W(1) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - WarpMineBit1X) < 13572 And Abs(W(4) - WarpMineBit1Y) < 349 Then
        
        WarpMineBit1YSpeed = (WarpMineBit1YSpeed * -1)
        WarpMineBit1Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - WarpMineBit1X) < 49 And Abs(W(6) - WarpMineBit1Y) < 6156 Then
        
        WarpMineBit1XSpeed = (WarpMineBit1XSpeed * -1)
        WarpMineBit1X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - WarpMineBit1X) < 349 And Abs(W(8) - WarpMineBit1Y) < 6156 Then
        
       WarpMineBit1XSpeed = (WarpMineBit1XSpeed * -1)
       WarpMineBit1X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(1) - WarpMineBit2X) < 13572 And Abs(W(2) - WarpMineBit2Y) < 49 Then
        
        WarpMineBit2YSpeed = (WarpMineBit2YSpeed * -1)
        WarpMineBit2Y = W(1) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(3) - WarpMineBit2X) < 13572 And Abs(W(4) - WarpMineBit2Y) < 349 Then
        
        WarpMineBit2YSpeed = (WarpMineBit2YSpeed * -1)
        WarpMineBit2Y = W(4) - 350
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(5) - WarpMineBit2X) < 49 And Abs(W(6) - WarpMineBit2Y) < 6156 Then
        
        WarpMineBit2XSpeed = (WarpMineBit2XSpeed * -1)
        WarpMineBit2X = W(5) + 50
        TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    If Abs(W(7) - WarpMineBit2X) < 349 And Abs(W(8) - WarpMineBit2Y) < 6156 Then
        
       WarpMineBit2XSpeed = (WarpMineBit2XSpeed * -1)
       WarpMineBit2X = W(7) - 350
       TotalPoints = TotalPoints + ((1) * PointMultiplier)
    
    End If
    
    'SAFETY MEASURES AGAINST OVERFLOW ERRORS
    If (TopSafetyCounterY - EnemyBit1Y) > -(50) Then
    
        EnemyBit1YSpeed = EnemyBit1YSpeed * -1
        EnemyBit1Y = EnemyBit1Y + 150
        
    End If
    
    If (BottomSafetyCounterY - EnemyBit1Y) < 25 Then
    
        EnemyBit1YSpeed = EnemyBit1YSpeed * -1
        EnemyBit1Y = EnemyBit1Y - 150
    
    End If
    
    If (LeftSafetyCounterX - EnemyBit1X) > -50 Then
    
        EnemyBit1XSpeed = EnemyBit1XSpeed * -1
        EnemyBit1X = EnemyBit1X + 150
    
    End If
    
    If (RightSafetyCounterX - EnemyBit1X) < 25 Then
    
        EnemyBit1XSpeed = EnemyBit1XSpeed * -1
        EnemyBit1X = EnemyBit1X - 150
    
    End If
    
    If (TopSafetyCounterY - EnemyBit2Y) > -(50) Then
    
        EnemyBit2YSpeed = EnemyBit2YSpeed * -1
        EnemyBit2Y = EnemyBit2Y + 150
        
    End If
    
    If (BottomSafetyCounterY - EnemyBit2Y) < 25 Then
    
        EnemyBit2YSpeed = EnemyBit2YSpeed * -1
        EnemyBit2Y = EnemyBit2Y - 150
    
    End If
    
    If (LeftSafetyCounterX - EnemyBit2X) > -50 Then
    
        EnemyBit2XSpeed = EnemyBit2XSpeed * -1
        EnemyBit2X = EnemyBit2X + 150
    
    End If
    
    If (RightSafetyCounterX - EnemyBit2X) < 25 Then
    
        EnemyBit2XSpeed = EnemyBit2XSpeed * -1
        EnemyBit2X = EnemyBit2X - 150
    
    End If
    
    If (TopSafetyCounterY - EnemyBit3Y) > -(50) Then
    
        EnemyBit3YSpeed = EnemyBit3YSpeed * -1
        EnemyBit3Y = EnemyBit3Y + 150
        
    End If
    
    If (BottomSafetyCounterY - EnemyBit3Y) < 25 Then
    
        EnemyBit3YSpeed = EnemyBit3YSpeed * -1
        EnemyBit3Y = EnemyBit3Y - 150
    
    End If
    
    If (LeftSafetyCounterX - EnemyBit3X) > -50 Then
    
        EnemyBit3XSpeed = EnemyBit3XSpeed * -1
        EnemyBit3X = EnemyBit3X + 150
    
    End If
    
    If (RightSafetyCounterX - EnemyBit3X) < 25 Then
    
        EnemyBit3XSpeed = EnemyBit3XSpeed * -1
        EnemyBit3X = EnemyBit3X - 150
    
    End If
    
    If (TopSafetyCounterY - EnemyBit4Y) > -(50) Then
    
        EnemyBit4YSpeed = EnemyBit4YSpeed * -1
        EnemyBit4Y = EnemyBit4Y + 150
        
    End If
    
    If (BottomSafetyCounterY - EnemyBit4Y) < 25 Then
    
        EnemyBit4YSpeed = EnemyBit4YSpeed * -1
        EnemyBit4Y = EnemyBit4Y - 150
    
    End If
    
    If (LeftSafetyCounterX - EnemyBit4X) > -50 Then
    
        EnemyBit4XSpeed = EnemyBit4XSpeed * -1
        EnemyBit4X = EnemyBit4X + 150
    
    End If
    
    If (RightSafetyCounterX - EnemyBit4X) < 25 Then
    
        EnemyBit4XSpeed = EnemyBit4XSpeed * -1
        EnemyBit4X = EnemyBit4X - 150
    
    End If
    
    If (TopSafetyCounterY - EnemyBit5Y) > -(50) Then
    
        EnemyBit5YSpeed = EnemyBit5YSpeed * -1
        EnemyBit5Y = EnemyBit5Y + 150
        
    End If
    
    If (BottomSafetyCounterY - EnemyBit5Y) < 25 Then
    
        EnemyBit5YSpeed = EnemyBit5YSpeed * -1
        EnemyBit5Y = EnemyBit5Y - 150
    
    End If
    
    If (LeftSafetyCounterX - EnemyBit5X) > -50 Then
    
        EnemyBit5XSpeed = EnemyBit5XSpeed * -1
        EnemyBit5X = EnemyBit5X + 150
    
    End If
    
    If (RightSafetyCounterX - EnemyBit5X) < 25 Then
    
        EnemyBit5XSpeed = EnemyBit5XSpeed * -1
        EnemyBit5X = EnemyBit5X - 150
    
    End If
    
    If (TopSafetyCounterY - EnemyBit6Y) > -(50) Then
    
        EnemyBit6YSpeed = EnemyBit6YSpeed * -1
        EnemyBit6Y = EnemyBit6Y + 150
        
    End If
    
    If (BottomSafetyCounterY - EnemyBit6Y) < 25 Then
    
        EnemyBit6YSpeed = EnemyBit6YSpeed * -1
        EnemyBit6Y = EnemyBit6Y - 150
    
    End If
    
    If (LeftSafetyCounterX - EnemyBit6X) > -50 Then
    
        EnemyBit6XSpeed = EnemyBit6XSpeed * -1
        EnemyBit6X = EnemyBit6X + 150
    
    End If
    
    If (RightSafetyCounterX - EnemyBit6X) < 25 Then
    
        EnemyBit6XSpeed = EnemyBit6XSpeed * -1
        EnemyBit6X = EnemyBit6X - 150
    
    End If
    
    If (TopSafetyCounterY - WarpMineBit1Y) > -(50) Then
    
        WarpMineBit1YSpeed = WarpMineBit1YSpeed * -1
        WarpMineBit1Y = WarpMineBit1Y + 150
        
    End If
    
    If (BottomSafetyCounterY - WarpMineBit1Y) < 25 Then
    
        WarpMineBit1YSpeed = WarpMineBit1YSpeed * -1
        WarpMineBit1Y = WarpMineBit1Y - 150
    
    End If
    
    If (LeftSafetyCounterX - WarpMineBit1X) > -50 Then
    
        WarpMineBit1XSpeed = WarpMineBit1XSpeed * -1
        WarpMineBit1X = WarpMineBit1X + 150
    
    End If
    
    If (RightSafetyCounterX - WarpMineBit1X) < 25 Then
    
        WarpMineBit1XSpeed = WarpMineBit1XSpeed * -1
        WarpMineBit1X = WarpMineBit1X - 150
    
    End If
    
    If (TopSafetyCounterY - WarpMineBit2Y) > -(50) Then
    
        WarpMineBit2YSpeed = WarpMineBit2YSpeed * -1
        WarpMineBit2Y = WarpMineBit2Y + 150
        
    End If
    
    If (BottomSafetyCounterY - WarpMineBit2Y) < 25 Then
    
        WarpMineBit2YSpeed = WarpMineBit2YSpeed * -1
        WarpMineBit2Y = WarpMineBit2Y - 150
    
    End If
    
    If (LeftSafetyCounterX - WarpMineBit2X) > -50 Then
    
        WarpMineBit2XSpeed = WarpMineBit2XSpeed * -1
        WarpMineBit2X = WarpMineBit2X + 150
    
    End If
    
    If (RightSafetyCounterX - WarpMineBit2X) < 25 Then
    
        WarpMineBit2XSpeed = WarpMineBit2XSpeed * -1
        WarpMineBit2X = WarpMineBit2X - 150
    
    End If
    
End Sub
Sub checkEnemyUserCollision()
    
    'Check Enemy-User Collisions
    If Abs(EnemyBit1X - UserBitX) < 300 And Abs(EnemyBit1Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(EnemyBit2X - UserBitX) < 300 And Abs(EnemyBit2Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(EnemyBit3X - UserBitX) < 300 And Abs(EnemyBit3Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(EnemyBit4X - UserBitX) < 300 And Abs(EnemyBit4Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(EnemyBit5X - UserBitX) < 300 And Abs(EnemyBit5Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(EnemyBit6X - UserBitX) < 300 And Abs(EnemyBit6Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(WarpMineBit1X - UserBitX) < 300 And Abs(WarpMineBit1Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

    If Abs(WarpMineBit2X - UserBitX) < 300 And Abs(WarpMineBit2Y - UserBitY) < 300 Then

        isEnemyHit = True

    End If

End Sub

