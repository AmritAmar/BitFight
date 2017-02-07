VERSION 5.00
Begin VB.Form frmSetUp 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitFight!"
   ClientHeight    =   8472
   ClientLeft      =   13896
   ClientTop       =   1620
   ClientWidth     =   8724
   Icon            =   "frmSetUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetUp.frx":058A
   ScaleHeight     =   8472
   ScaleWidth      =   8724
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Colour Changing"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   8172
      Left            =   6120
      TabIndex        =   12
      Top             =   120
      Width           =   2412
      Begin VB.CommandButton cmdCancelColour 
         Caption         =   "Cancel"
         Height          =   372
         Left            =   240
         TabIndex        =   25
         Top             =   6960
         Width           =   1932
      End
      Begin VB.CommandButton cmdApplyColour 
         Caption         =   "Apply"
         Height          =   372
         Left            =   240
         TabIndex        =   17
         Top             =   7440
         Width           =   1932
      End
      Begin VB.VScrollBar VSBRed 
         Height          =   3612
         LargeChange     =   10
         Left            =   360
         Max             =   255
         SmallChange     =   5
         TabIndex        =   15
         Top             =   2040
         Width           =   492
      End
      Begin VB.VScrollBar VSBBlue 
         Height          =   3612
         LargeChange     =   10
         Left            =   1560
         Max             =   255
         SmallChange     =   5
         TabIndex        =   14
         Top             =   2040
         Width           =   492
      End
      Begin VB.VScrollBar VSBGreen 
         Height          =   3612
         LargeChange     =   10
         Left            =   960
         Max             =   255
         SmallChange     =   5
         TabIndex        =   13
         Top             =   2040
         Width           =   492
      End
      Begin VB.Label lblBlueValue 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   1560
         TabIndex        =   28
         Top             =   5760
         Width           =   492
      End
      Begin VB.Label lblGreenValue 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   960
         TabIndex        =   27
         Top             =   5760
         Width           =   492
      End
      Begin VB.Label lblRedValue 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   360
         TabIndex        =   26
         Top             =   5760
         Width           =   492
      End
      Begin VB.Label lblColourEdit 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Colour Edit ##"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   372
         Left            =   240
         TabIndex        =   24
         Top             =   6240
         Width           =   1932
      End
      Begin VB.Label OldColour 
         BorderStyle     =   1  'Fixed Single
         Height          =   1332
         Left            =   480
         TabIndex        =   23
         Top             =   480
         Width           =   732
      End
      Begin VB.Label NewColour 
         BorderStyle     =   1  'Fixed Single
         Height          =   1332
         Left            =   1200
         TabIndex        =   16
         Top             =   480
         Width           =   732
      End
   End
   Begin VB.Timer timColours 
      Interval        =   20
      Left            =   5760
      Top             =   7680
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Start Game!"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1332
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   5532
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start BitFight!"
         Height          =   612
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   2772
      End
   End
   Begin VB.Frame fraOptions1 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5292
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5532
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Bit Colors"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3972
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   5292
         Begin VB.CheckBox chkCircularWarpMine 
            BackColor       =   &H00000000&
            Caption         =   "Circular Warp Mines"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   252
            Left            =   2520
            TabIndex        =   31
            Top             =   3240
            Width           =   2172
         End
         Begin VB.CheckBox chkCirclularPowerUp 
            BackColor       =   &H00000000&
            Caption         =   "Circular Power Ups"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   252
            Left            =   240
            TabIndex        =   30
            Top             =   3600
            Width           =   2052
         End
         Begin VB.CheckBox chkSolid 
            BackColor       =   &H00000000&
            Caption         =   "Solid Colours"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   252
            Left            =   240
            TabIndex        =   29
            Top             =   3240
            Width           =   1932
         End
         Begin VB.CommandButton cmdWarpMinesChange 
            Caption         =   "Change Colour"
            Height          =   372
            Left            =   2520
            TabIndex        =   22
            Top             =   2760
            Width           =   2532
         End
         Begin VB.CommandButton cmdEnemyBitChange 
            Caption         =   "Change Colour"
            Height          =   372
            Left            =   2520
            TabIndex        =   21
            Top             =   2160
            Width           =   2532
         End
         Begin VB.CommandButton cmdDistortionChange 
            Caption         =   "Change Colour"
            Height          =   372
            Left            =   2520
            TabIndex        =   20
            Top             =   1560
            Width           =   2532
         End
         Begin VB.CommandButton cmdPowerUpChange 
            Caption         =   "Change Colour"
            Height          =   372
            Left            =   2520
            TabIndex        =   19
            Top             =   960
            Width           =   2532
         End
         Begin VB.CommandButton cmdUserBitChange 
            Caption         =   "Change Colour"
            Height          =   372
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   2532
         End
         Begin VB.Shape UserBitColour 
            BackColor       =   &H00FF0000&
            BorderColor     =   &H00FF0000&
            FillColor       =   &H00FF0000&
            FillStyle       =   6  'Cross
            Height          =   300
            Left            =   240
            Shape           =   5  'Rounded Square
            Top             =   360
            Width           =   300
         End
         Begin VB.Shape PowerUpColour 
            BackColor       =   &H0000FFFF&
            BorderColor     =   &H0000FFFF&
            FillColor       =   &H0000FFFF&
            FillStyle       =   3  'Vertical Line
            Height          =   300
            Left            =   240
            Shape           =   5  'Rounded Square
            Top             =   960
            Width           =   300
         End
         Begin VB.Shape DistortionBitColour 
            BackColor       =   &H00FFFF00&
            BorderColor     =   &H00FFFF00&
            FillColor       =   &H00FFFF00&
            FillStyle       =   6  'Cross
            Height          =   300
            Left            =   240
            Shape           =   1  'Square
            Top             =   1560
            Width           =   300
         End
         Begin VB.Shape EnemyBitColour 
            BackColor       =   &H000000FF&
            BorderColor     =   &H000000FF&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H000000FF&
            FillStyle       =   4  'Upward Diagonal
            Height          =   300
            Left            =   240
            Shape           =   1  'Square
            Top             =   2160
            Width           =   300
         End
         Begin VB.Shape WarpMineColour 
            BackColor       =   &H000080FF&
            BorderColor     =   &H000080FF&
            FillColor       =   &H000080FF&
            FillStyle       =   7  'Diagonal Cross
            Height          =   300
            Left            =   240
            Shape           =   1  'Square
            Top             =   2760
            Width           =   300
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "Distortion"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   372
            Left            =   720
            TabIndex        =   11
            Top             =   1560
            Width           =   1692
         End
         Begin VB.Label Label10 
            BackColor       =   &H00000000&
            Caption         =   "Warp Mines"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   372
            Left            =   720
            TabIndex        =   8
            Top             =   2760
            Width           =   1452
         End
         Begin VB.Label Label9 
            BackColor       =   &H00000000&
            Caption         =   "Power Ups"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   372
            Left            =   720
            TabIndex        =   7
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Label8 
            BackColor       =   &H00000000&
            Caption         =   "Enemy Bit"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   372
            Left            =   720
            TabIndex        =   6
            Top             =   2160
            Width           =   1212
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   "Your Bit"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   372
            Left            =   720
            TabIndex        =   5
            Top             =   360
            Width           =   972
         End
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   372
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "Bit Fighter"
         Top             =   480
         Width           =   3852
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   492
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.Image imgTitle 
      BorderStyle     =   1  'Fixed Single
      Height          =   1332
      Left            =   120
      Picture         =   "frmSetUp.frx":7A0EC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5532
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   5880
      X2              =   5880
      Y1              =   120
      Y2              =   8280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Beta Release - Version 1.0"
      ForeColor       =   &H0000FF00&
      Height          =   252
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   1932
   End
End
Attribute VB_Name = "frmSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApplyColour_Click()
    
    frmSetUp.Width = 5900
    
    If ColourEditor = 1 Then
    
        UserBitColour.BackColor = NewColour.BackColor
        UserBitColour.FillColor = NewColour.BackColor
        UserBitColour.BorderColor = NewColour.BackColor
        
    ElseIf ColourEditor = 2 Then
    
        PowerUpColour.BackColor = NewColour.BackColor
        PowerUpColour.FillColor = NewColour.BackColor
        PowerUpColour.BorderColor = NewColour.BackColor
        
    ElseIf ColourEditor = 3 Then
    
        DistortionBitColour.BackColor = NewColour.BackColor
        DistortionBitColour.FillColor = NewColour.BackColor
        DistortionBitColour.BorderColor = NewColour.BackColor
        
    ElseIf ColourEditor = 4 Then
    
        EnemyBitColour.BackColor = NewColour.BackColor
        EnemyBitColour.FillColor = NewColour.BackColor
        EnemyBitColour.BorderColor = NewColour.BackColor
        
    ElseIf ColourEditor = 5 Then
        
        WarpMineColour.BackColor = NewColour.BackColor
        WarpMineColour.FillColor = NewColour.BackColor
        WarpMineColour.BorderColor = NewColour.BackColor
        
    End If

End Sub

Private Sub cmdCancelColour_Click()

    frmSetUp.Width = 5900

End Sub

Private Sub cmdDistortionChange_Click()

    frmSetUp.Width = 8808 'Sets New Width
    ColourEditor = 3 'Sets Current Editor to 3 (Distortion BIT)
    OldColour.BackColor = DistortionBitColour.BackColor 'Sets Old Back Color
    lblColourEdit.Caption = "Distortion Colour" 'Sets Caption for UI

End Sub

Private Sub cmdEnemyBitChange_Click()

    frmSetUp.Width = 8808 'Sets New Width
    ColourEditor = 4 'Sets Current Editor to 4 (Enemy BIT)
    OldColour.BackColor = EnemyBitColour.BackColor 'Sets Old Back Color
    lblColourEdit.Caption = "Enemy Bit" 'Sets Caption for UI
    
End Sub

Private Sub cmdPowerUpChange_Click()

    frmSetUp.Width = 8808 'Sets New Width
    ColourEditor = 2 'Sets Current Editor to 2 (Power Up BIT)
    OldColour.BackColor = PowerUpColour.BackColor 'Sets Old Back Color
    lblColourEdit.Caption = "Power Up Bit" 'Sets Caption for UI

End Sub

Private Sub cmdStart_Click()

    UserBitBGColour = UserBitColour.BackColor
    EnemyBitBGColour = EnemyBitColour.BackColor
    PowerUpBGColour = PowerUpColour.BackColor
    WarpMineBGColour = WarpMineColour.BackColor
    DistortionBGColour = DistortionBitColour.BackColor
    PlayerName = txtName.Text
    
    frmMainMenu.Show
    
    Unload Me

End Sub

Private Sub cmdUserBitChange_Click()

    frmSetUp.Width = 8808 'Sets New Width
    ColourEditor = 1 'Sets Current Editor to 1 (USER BIT)
    OldColour.BackColor = UserBitColour.BackColor 'Sets Old Back Color
    lblColourEdit.Caption = "Your Bit" 'Sets Caption for UI

End Sub

Private Sub cmdWarpMinesChange_Click()

    frmSetUp.Width = 8808 'Sets New Width
    ColourEditor = 5 'Sets Current Editor to 5 (Warp Mine BIT)
    OldColour.BackColor = WarpMineColour.BackColor 'Sets Old Back Color
    lblColourEdit.Caption = "Warp Mine Bit" 'Sets Caption for UI
   
End Sub

Private Sub Form_Load()

    cmdStart.TabIndex = 1
    frmSetUp.Width = 5900

End Sub

Private Sub timColours_Timer()

    If chkSolid <> 0 Then
    
        isSolidColours = True
        
    Else: isSolidColours = False
    
    End If
    
    If chkCirclularPowerUp <> 0 Then
    
        isCircularPowerUp = True
        
    Else: isCircularPowerUp = False
    
    End If
    
    If isCircularPowerUp = True Then
    
        PowerUpColour.Shape = ShapeConstants.vbShapeCircle
    
    Else: PowerUpColour.Shape = ShapeConstants.vbShapeRoundedSquare
    
    End If
    
    If chkCircularWarpMine <> 0 Then
    
        isCircularWarpMine = True
        
    Else: isCircularWarpMine = False
    
    End If
    
    If isCircularWarpMine = True Then
    
        WarpMineColour.Shape = ShapeConstants.vbShapeCircle
    
    Else: WarpMineColour.Shape = ShapeConstants.vbShapeSquare
    End If
    
    If isSolidColours = True Then
    
        UserBitColour.BackStyle = 1
        PowerUpColour.BackStyle = 1
        EnemyBitColour.BackStyle = 1
        WarpMineColour.BackStyle = 1
        
    Else
    
        UserBitColour.BackStyle = 0
        PowerUpColour.BackStyle = 0
        DistortionBitColour.BackStyle = 0
        EnemyBitColour.BackStyle = 0
        WarpMineColour.BackStyle = 0
        
    End If

    NewColour.BackColor = RGB(VSBRed, VSBGreen, VSBBlue)

    lblRedValue = VSBRed.Value
    lblGreenValue = VSBGreen.Value
    lblBlueValue = VSBBlue.Value

End Sub
