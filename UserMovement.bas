Attribute VB_Name = "UserMovement"
'-----------------------------------------------------------------
'                         USER ACTIONS
'           All User Action Code happens in this Module
'-----------------------------------------------------------------


Sub checkKey(KeyCode As Integer, Shift As Integer) 'Checks key pressed and does action
    
    If isExploded = False Then
    
        If KeyCode = KeyCodeConstants.vbKeyRight Or KeyCode = KeyCodeConstants.vbKeyD Then
        
            UserBitXSpeed = UserBitXSpeed + SpeedIncrementation
            LeftFire = True
        
        ElseIf KeyCode = KeyCodeConstants.vbKeyLeft Or KeyCode = KeyCodeConstants.vbKeyA Then
        
            UserBitXSpeed = UserBitXSpeed - SpeedIncrementation
            RightFire = True
        
        ElseIf KeyCode = KeyCodeConstants.vbKeyUp Or KeyCode = KeyCodeConstants.vbKeyW Then
        
            UserBitYSpeed = UserBitYSpeed - SpeedIncrementation
            DownFire = True
        
        ElseIf KeyCode = KeyCodeConstants.vbKeyDown Or KeyCode = KeyCodeConstants.vbKeyS Then
        
            UserBitYSpeed = UserBitYSpeed + SpeedIncrementation
            TopFire = True
        
        ElseIf KeyCode = KeyCodeConstants.vbKeySpace Then
        
            If isSlowMotion = True Then
        
                isSlowMotion = False
        
            Else
        
                isSlowMotion = True
        
            End If
        
        ElseIf KeyCode = KeyCodeConstants.vbKeyQ Or KeyCode = KeyCodeConstants.vbKeyN Then
        
            If DistortionValue > 0 Then
        
                isDistortion = True
        
            Else
        
                isDistortion = False
        
            End If
        
        ElseIf KeyCode = KeyCodeConstants.vbKeyE Or KeyCode = KeyCodeConstants.vbKeyB Then
        
            If isEMPOn = False And EMPValue = 200 Then
        
                isEMPOn = True
                isEMPTrigerred = True
        
            End If
         
        ElseIf KeyCode = KeyCodeConstants.vbKeyJ And Shift = 1 Then
        
            TotalPoints = TotalPoints + 100
            PointMultiplier = PointMultiplier + 1
        
        End If
        
    End If

End Sub

Sub checkKeyUp(KeyCode As Integer, Shift As Integer) 'Checks Keys Up and does action

    If KeyCode = KeyCodeConstants.vbKeyQ Or KeyCode = KeyCodeConstants.vbKeyN Then
    
        isDistortion = False
    
    End If

End Sub

Sub UserBitMovement(UserBitX As Integer, UserBitY As Integer, UserBitXSpeed As Integer, UserBitYSpeed As Integer) 'Set Movement and Mapping
    
    If isExploded = False Then
    
        'Set User Bit to Move
        UserBitX = UserBitX + UserBitXSpeed
        UserBitY = UserBitY + UserBitYSpeed
    
        'Map out Directional Speeds
        UserBitUpSpeed = UserBitYSpeed * -1
        UserBitDownSpeed = UserBitYSpeed
        UserBitLeftSpeed = UserBitXSpeed * -1
        UserBitRightSpeed = UserBitXSpeed
        
    End If
    
End Sub

Sub DistortionBitMovement(UserBitX As Integer, UserBitY As Integer, DistortionBitX As Integer, DistortionBitY As Integer) 'Set Distortion Mapping

    DistortionBitX = (UserBitX)
    DistortionBitY = (UserBitY)

End Sub

Sub PowerUpBitMovement()
    
    If isExploded = False Then
    
        'These Create Random (X,Y) Positions for the Power Up
        PowerUpBitX = CInt(Int((12600 - 300 + 1) * Rnd() + 300))
        PowerUpBitY = CInt(Int((5500 - 300 + 1) * Rnd() + 300))
    
        PowerUpBitNumber = CInt(Int((30 - 1 + 1) * Rnd() + 1))

    End If

End Sub

