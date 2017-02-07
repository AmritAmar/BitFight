VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSurvivalEndless 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitFight! - Survival - Endless"
   ClientHeight    =   8376
   ClientLeft      =   960
   ClientTop       =   1692
   ClientWidth     =   17028
   Icon            =   "BitFight! - Survival - Endless.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BitFight! - Survival - Endless.frx":058A
   ScaleHeight     =   8376
   ScaleWidth      =   17028
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRestart 
      BackColor       =   &H00000000&
      Height          =   2892
      Left            =   8760
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   4812
      Begin VB.CommandButton cmdMainMenu 
         Caption         =   "Go To Main Menu!"
         Height          =   612
         Left            =   2520
         TabIndex        =   31
         Top             =   1920
         Width           =   2052
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Play Again!"
         Height          =   612
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   2052
      End
      Begin VB.Label lblGameOver 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Game Over. Your Bit has been Obliterated because you collided with your Enemy. You have lost the battle."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1332
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   4332
      End
   End
   Begin VB.Timer timEndlessPosCheck 
      Interval        =   1
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer timWarpMineBit 
      Interval        =   2000
      Left            =   2880
      Top             =   120
   End
   Begin VB.Timer timUserMotion 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.Frame frmHolderInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   2052
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   16932
      Begin VB.Frame frmBitInfo 
         BackColor       =   &H00000000&
         Caption         =   "Bit Information"
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
         Height          =   1932
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   16860
         Begin VB.TextBox txtUpSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   552
            Left            =   14280
            TabIndex        =   8
            Top             =   480
            Width           =   1092
         End
         Begin VB.TextBox txtLeftSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   552
            Left            =   13080
            TabIndex        =   7
            Top             =   1080
            Width           =   1092
         End
         Begin VB.TextBox txtDownSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   552
            Left            =   14280
            TabIndex        =   6
            Top             =   1080
            Width           =   1092
         End
         Begin VB.TextBox txtRightSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   552
            Left            =   15480
            TabIndex        =   5
            Top             =   1080
            Width           =   1092
         End
         Begin MSComctlLib.ProgressBar prbDistortion 
            Height          =   372
            Left            =   6480
            TabIndex        =   9
            Top             =   840
            Width           =   5292
            _ExtentX        =   9335
            _ExtentY        =   656
            _Version        =   393216
            Appearance      =   1
            Max             =   200
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar prbSlowMotion 
            Height          =   372
            Left            =   6480
            TabIndex        =   10
            Top             =   360
            Width           =   5292
            _ExtentX        =   9335
            _ExtentY        =   656
            _Version        =   393216
            Appearance      =   1
            Max             =   200
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar prbEMP 
            Height          =   372
            Left            =   6480
            TabIndex        =   11
            Top             =   1320
            Width           =   5292
            _ExtentX        =   9335
            _ExtentY        =   656
            _Version        =   393216
            Appearance      =   1
            Max             =   200
            Scrolling       =   1
         End
         Begin VB.Label lblSlowMotion_OffOn 
            BackColor       =   &H00000000&
            Caption         =   "Y/N"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   492
            Left            =   11880
            TabIndex        =   24
            Top             =   360
            Width           =   972
         End
         Begin VB.Label lblDistortion_OffOn 
            BackColor       =   &H00000000&
            Caption         =   "Y/N"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   11880
            TabIndex        =   23
            Top             =   840
            Width           =   972
         End
         Begin VB.Label lblSlowMotion_Value 
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   5400
            TabIndex        =   22
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblSlowMotion_Name 
            BackColor       =   &H00000000&
            Caption         =   "Slow Motion:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   492
            Left            =   3240
            TabIndex        =   21
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label lblDistortion_Value 
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   5400
            TabIndex        =   20
            Top             =   840
            Width           =   1092
         End
         Begin VB.Label lblDistortion_Name 
            BackColor       =   &H00000000&
            Caption         =   "Distortion:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   3600
            TabIndex        =   19
            Top             =   840
            Width           =   1692
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   492
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   3012
         End
         Begin VB.Label lblName 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NAME"
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
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   2892
         End
         Begin VB.Label lblPointMultiplier_Name 
            BackColor       =   &H00000000&
            Caption         =   "Point Multiplier:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   2652
         End
         Begin VB.Label lblPointMultiplier_Value 
            BackColor       =   &H00000000&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   372
            Left            =   2760
            TabIndex        =   15
            Top             =   960
            Width           =   612
         End
         Begin VB.Label lblEMP_Name 
            BackColor       =   &H00000000&
            Caption         =   "EMP:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   4440
            TabIndex        =   14
            Top             =   1320
            Width           =   852
         End
         Begin VB.Label lblEMP_Value 
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   5400
            TabIndex        =   13
            Top             =   1320
            Width           =   1092
         End
         Begin VB.Label lblEMP_OffOn 
            BackColor       =   &H00000000&
            Caption         =   "Ready"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   16.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   11880
            TabIndex        =   12
            Top             =   1320
            Width           =   1092
         End
      End
   End
   Begin VB.Frame frmHolderLog 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6372
      Left            =   13800
      TabIndex        =   0
      Top             =   0
      Width           =   3252
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "System Log"
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
         Height          =   6252
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3132
         Begin VB.TextBox txtSystemLog 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   5772
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   2892
         End
      End
   End
   Begin VB.Timer timEnemyMotion 
      Interval        =   1
      Left            =   360
      Top             =   120
   End
   Begin VB.Timer timSurvivalEndless 
      Interval        =   1
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer timGameMechanics 
      Interval        =   1
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer timPowerUp 
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer timRegeneration 
      Interval        =   2500
      Left            =   1800
      Top             =   120
   End
   Begin VB.Timer timPowerUpResultant 
      Interval        =   5000
      Left            =   2160
      Top             =   120
   End
   Begin VB.Timer timEMP 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2520
      Top             =   120
   End
   Begin VB.Image Explosion1 
      Height          =   996
      Left            =   5520
      Picture         =   "BitFight! - Survival - Endless.frx":7A0EC
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.Shape WarpMineBit2 
      BorderColor     =   &H00FF00FF&
      FillColor       =   &H00FF00FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   300
      Left            =   5400
      Shape           =   1  'Square
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape EnemyBit6 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H000000FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   300
      Left            =   6840
      Shape           =   1  'Square
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape EnemyBit5 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   300
      Left            =   5760
      Shape           =   1  'Square
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape PowerUpBit 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   3  'Vertical Line
      Height          =   300
      Left            =   7200
      Shape           =   5  'Rounded Square
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape DistortionBit 
      BackColor       =   &H000080FF&
      BorderColor     =   &H000080FF&
      BorderStyle     =   3  'Dot
      BorderWidth     =   3
      FillColor       =   &H000080FF&
      FillStyle       =   6  'Cross
      Height          =   300
      Left            =   7560
      Shape           =   1  'Square
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape EnemyBit3 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   300
      Left            =   5760
      Shape           =   1  'Square
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape WarpMineBit1 
      BorderColor     =   &H00FF00FF&
      FillColor       =   &H00FF00FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   300
      Left            =   5400
      Shape           =   1  'Square
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape EnemyBit1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H000000FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   300
      Left            =   6480
      Shape           =   1  'Square
      Top             =   1920
      Width           =   300
   End
   Begin VB.Image LeftFire1 
      Height          =   300
      Left            =   6720
      Picture         =   "BitFight! - Survival - Endless.frx":7A583
      Stretch         =   -1  'True
      Top             =   960
      Width           =   300
   End
   Begin VB.Image DownFire1 
      Height          =   300
      Left            =   7080
      Picture         =   "BitFight! - Survival - Endless.frx":7A9C7
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image RightFire1 
      Height          =   300
      Left            =   7440
      Picture         =   "BitFight! - Survival - Endless.frx":7ADAF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   300
   End
   Begin VB.Shape UserBit 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      FillColor       =   &H00FF0000&
      FillStyle       =   6  'Cross
      Height          =   300
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   300
   End
   Begin VB.Shape EnemyBit2 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   300
      Left            =   6120
      Shape           =   1  'Square
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image TopFire1 
      Height          =   300
      Left            =   7080
      Picture         =   "BitFight! - Survival - Endless.frx":7B1A6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   300
   End
   Begin VB.Shape EnemyBit4 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H000000FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   300
      Left            =   6840
      Shape           =   1  'Square
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label TopCounterMeasure 
      BackColor       =   &H00FFFFFF&
      Height          =   96
      Left            =   360
      TabIndex        =   28
      Top             =   0
      Width           =   96
   End
   Begin VB.Label LeftCounterMeasure 
      Height          =   96
      Left            =   240
      TabIndex        =   27
      Top             =   0
      Width           =   96
   End
   Begin VB.Label RightCounterMeasure 
      Height          =   96
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Width           =   96
   End
   Begin VB.Label BottomCounterMeasure 
      Height          =   96
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   96
   End
End
Attribute VB_Name = "frmSurvivalEndless"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMainMenu_Click()

    Unload Me
    frmMainMenu.Show

End Sub

Private Sub cmdReset_Click()

    If cmdReset.Caption = "Go to ZomBits!" Then
    
        MsgBox ("Loading: ZomBits... Good Luck " & PlayerName & "!")
        Unload Me
        frmZomBits.Show
        
    Else
        
        MsgBox ("Loading: Survival Endless... Good Luck " & PlayerName & "!")
        Unload Me
        frmSurvivalEndless.Show
    
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Call Sub-Routine to Check Keys Pressed and hence do actions
    Call checkKey(KeyCode, Shift)

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'Call Sub-Routine to Check Keys Up and hence do actions
    Call checkKeyUp(KeyCode, Shift)

End Sub

Private Sub Form_Load()
    
    Randomize
    'Change The Text Box to Show the Following Intructions in the Side log
    txtSystemLog.Text = " - LOG INITIALIZED - " _
                            + vbNewLine + vbNewLine _
                            + ">>> BitFight Routine Started!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Difficulty: ENDLESS" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> OBJECTIVE: Survive. Just Survive..." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You will get 1 Point for every Enemy-Wall Collision, and Point Multipliers will be added for getting consecutive Power Ups that appear on the Arena..." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Good Luck " + PlayerName + "!"
    
    'Call the Start up Routine
    Call StartUpSurvivalEndless
    
    'Randomize Function
    Randomize
    
    'General Things to do
    prbSlowMotion.Value = SlowMotionValue 'Set Slow Mo Value
    'timPowerUp.Interval = 1000
    
    'Colour Allocation
    UserBit.FillColor = UserBitBGColour
    EnemyBit1.FillColor = EnemyBitBGColour
    EnemyBit2.FillColor = EnemyBitBGColour
    EnemyBit3.FillColor = EnemyBitBGColour
    EnemyBit4.FillColor = EnemyBitBGColour
    EnemyBit5.FillColor = EnemyBitBGColour
    EnemyBit6.FillColor = EnemyBitBGColour
    WarpMineBit1.FillColor = WarpMineBGColour
    WarpMineBit2.FillColor = WarpMineBGColour
    PowerUpBit.FillColor = PowerUpBGColour
    DistortionBit.FillColor = DistortionBGColour
    UserBit.BackColor = UserBitBGColour
    EnemyBit1.BackColor = EnemyBitBGColour
    EnemyBit2.BackColor = EnemyBitBGColour
    EnemyBit3.BackColor = EnemyBitBGColour
    EnemyBit4.BackColor = EnemyBitBGColour
    EnemyBit5.BackColor = EnemyBitBGColour
    EnemyBit6.BackColor = EnemyBitBGColour
    WarpMineBit1.BackColor = WarpMineBGColour
    WarpMineBit2.BackColor = WarpMineBGColour
    PowerUpBit.BackColor = PowerUpBGColour
    DistortionBit.BackColor = DistortionBGColour
    UserBit.BorderColor = UserBitBGColour
    EnemyBit1.BorderColor = EnemyBitBGColour
    EnemyBit2.BorderColor = EnemyBitBGColour
    EnemyBit3.BorderColor = EnemyBitBGColour
    EnemyBit4.BorderColor = EnemyBitBGColour
    EnemyBit5.BorderColor = EnemyBitBGColour
    EnemyBit6.BorderColor = EnemyBitBGColour
    WarpMineBit1.BorderColor = WarpMineBGColour
    WarpMineBit2.BorderColor = WarpMineBGColour
    PowerUpBit.BorderColor = PowerUpBGColour
    DistortionBit.BorderColor = DistortionBGColour
    
    If isSolidColours = True Then
    
        UserBit.BackStyle = 1
        EnemyBit1.BackStyle = 1
        EnemyBit2.BackStyle = 1
        EnemyBit3.BackStyle = 1
        EnemyBit4.BackStyle = 1
        EnemyBit5.BackStyle = 1
        EnemyBit6.BackStyle = 1
        WarpMineBit1.BackStyle = 1
        WarpMineBit2.BackStyle = 1
        PowerUpBit.BackStyle = 1
        
    End If
    
    If isCircularPowerUp = True Then
    
        PowerUpBit.Shape = ShapeConstants.vbShapeCircle
        
    End If
    
    If isCircularWarpMine = True Then
    
        WarpMineBit1.Shape = ShapeConstants.vbShapeCircle
        WarpMineBit2.Shape = ShapeConstants.vbShapeCircle
        
    End If
    
    BottomCounterMeasure.BackColor = ColorConstants.vbBlack
    TopCounterMeasure.BackColor = ColorConstants.vbBlack
    LeftCounterMeasure.BackColor = ColorConstants.vbBlack
    RightCounterMeasure.BackColor = ColorConstants.vbBlack
    
    'Name Allocation
    lblName.Caption = PlayerName
    
    'Set Position of Safety Counters
    TopCounterMeasure.Top = TopSafetyCounterY
    TopCounterMeasure.Left = TopSafetyCounterX
    LeftCounterMeasure.Top = LeftSafetyCounterY
    LeftCounterMeasure.Left = LeftSafetyCounterX
    RightCounterMeasure.Top = RightSafetyCounterY
    RightCounterMeasure.Left = RightSafetyCounterX
    BottomCounterMeasure.Top = BottomSafetyCounterY
    BottomCounterMeasure.Left = BottomSafetyCounterX
    
    'TIMERS ENABLED
    timUserMotion.Enabled = True
    timEnemyMotion.Enabled = True
    timSurvivalEndless.Enabled = True
    timGameMechanics.Enabled = True
    timPowerUp.Enabled = True

End Sub

Private Sub Form_Paint()

    'Draw all the Walls
    Line (W(1), W(2))-(W(1) + 13572, W(2) + 15), RGB(0, 255, 0), BF
    Line (W(3), W(4))-(W(3) + 13572, W(4) + 15), RGB(0, 255, 0), BF
    Line (W(5), W(6))-(W(5) + 15, W(6) + 6156), RGB(0, 255, 0), BF
    Line (W(7), W(8))-(W(7) + 15, W(8) + 6156), RGB(0, 255, 0), BF
    
End Sub

Private Sub timEMP_Timer() 'The EMP Timer

    If isEMPOn = True Then  'Checks if EMP is On
    
        If EMPValue <> 200 Then 'Checks the Value of EMP and Increments it
        
            EMPValue = EMPValue + 0.25 'Incrementation of EMP Value
            
        End If 'End Checking
        
        If EMPValue > 50 Then 'Checks if Enemies are allowed to go free
        
            isEMPOn = False 'Sets the EMP to off permanently
            
        End If 'End Checking
        
    Else 'Else if EMP is false

        If EMPValue <> 200 Then 'Still Increments until the Value is completely filled
            
            EMPValue = EMPValue + 0.25
                
        Else
        
            'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Your EMP is charged up!"
                            
            lblEMP_OffOn.Caption = "Ready" 'Updates UI
            timEMP.Enabled = False 'Disables Timer
            
        End If
    
    End If
    
End Sub

Private Sub timEndlessPosCheck_Timer() 'Keeps checking and Updating the Positions if not enabled

    If EndlessSurvivalProgression = 0 Then
    
        EnemyBit2X = -300
        EnemyBit2Y = -300
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
        
    ElseIf EndlessSurvivalProgression = 1 Then
    
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
        
    ElseIf EndlessSurvivalProgression = 2 Then
    
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
        
    ElseIf EndlessSurvivalProgression = 3 Then
    
        EnemyBit5X = -300
        EnemyBit5Y = -300
        EnemyBit6X = -300
        EnemyBit6Y = -300
        WarpMineBit1X = -300
        WarpMineBit1Y = -300
        WarpMineBit2X = -300
        WarpMineBit2Y = -300
        
    ElseIf EndlessSurvivalProgression = 4 Then
    
        EnemyBit5X = -300
        EnemyBit5Y = -300
        EnemyBit6X = -300
        EnemyBit6Y = -300
        WarpMineBit2X = -300
        WarpMineBit2Y = -300
        
    ElseIf EndlessSurvivalProgression = 5 Then
    
        EnemyBit6X = -300
        EnemyBit6Y = -300
        WarpMineBit2X = -300
        WarpMineBit2Y = -300
        
    ElseIf EndlessSurvivalProgression = 6 Then

        EnemyBit6X = -300
        EnemyBit6Y = -300
        
    Else
    
    End If
    
End Sub

Private Sub timEnemyMotion_Timer() 'Makes Enemies move, depending on Points and Progression

    'Enemy Bit 1
    EnemyBit1.Visible = True
    EnemyBit1.Left = EnemyBit1X
    EnemyBit1.Top = EnemyBit1Y
    
    'Movement
    Call EnemyBitMovement(EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
    
    If TotalPoints > 100 Then
        
        If EndlessSurvivalProgression = 0 Then
        
            'Text on Log to inform Player
            txtSystemLog.Text = txtSystemLog.Text _
                        + vbNewLine _
                        + vbNewLine _
                        + ">>> New Enemy Added!"
            EndlessSurvivalProgression = 1
            EnemyBit2X = 12000
            EnemyBit2Y = 2925
            
        End If
        
        EnemyBit2.Visible = True
        EnemyBit2.Left = EnemyBit2X
        EnemyBit2.Top = EnemyBit2Y
        Call EnemyBitMovement(EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
        
        If TotalPoints > 250 Then
        
            If EndlessSurvivalProgression = 1 Then
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> New Enemy Added!"
                EndlessSurvivalProgression = 2
                EnemyBit3X = 6635
                EnemyBit3Y = 500
                
            End If
            
            EnemyBit3.Visible = True
            EnemyBit3.Left = EnemyBit3X
            EnemyBit3.Top = EnemyBit3Y
            Call EnemyBitMovement(EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
        
            If TotalPoints > 650 Then
            
                If EndlessSurvivalProgression = 2 Then
                    
                    'Text on Log to inform Player
                    txtSystemLog.Text = txtSystemLog.Text _
                                + vbNewLine _
                                + vbNewLine _
                                + ">>> New Enemy Added!"
                    EndlessSurvivalProgression = 3
                    EnemyBit4X = 6635
                    EnemyBit4Y = 2925
                    
                End If
                
                EnemyBit4.Visible = True
                EnemyBit4.Left = EnemyBit4X
                EnemyBit4.Top = EnemyBit4Y
                Call EnemyBitMovement(EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
                
                If TotalPoints > 1000 Then
                
                    If EndlessSurvivalProgression = 4 Then
                    
                        'Text on Log to inform Player
                        txtSystemLog.Text = txtSystemLog.Text _
                                    + vbNewLine _
                                    + vbNewLine _
                                    + ">>> New Enemy Added!"
                        EndlessSurvivalProgression = 5
                        EnemyBit5X = 1000
                        EnemyBit5Y = 2925
                    
                    End If
                    
                    EnemyBit5.Visible = True
                    EnemyBit5.Left = EnemyBit5X
                    EnemyBit5.Top = EnemyBit5Y
                    Call EnemyBitMovement(EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
                        
                    If TotalPoints > 1200 Then
                        
                        If EndlessSurvivalProgression = 6 Then
                        
                            'Text on Log to inform Player
                            txtSystemLog.Text = txtSystemLog.Text _
                                        + vbNewLine _
                                        + vbNewLine _
                                        + ">>> New Enemy Added!"
                            EndlessSurvivalProgression = 7
                            EnemyBit6X = 12000
                            EnemyBit6Y = 2925
                            
                        End If
                        
                        EnemyBit6.Visible = True
                        EnemyBit6.Left = EnemyBit6X
                        EnemyBit6.Top = EnemyBit6Y
                        Call EnemyBitMovement(EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                    
                    End If
                    
                End If
                
            End If
            
        End If
    
    End If
    
    Call checkEnemyWallCollision 'Checks Wall Collisions
    
End Sub

Private Sub timgamemechanics_Timer()
    
    'Updates Position of Power Up
    PowerUpBit.Left = PowerUpBitX
    PowerUpBit.Top = PowerUpBitY

    'Update UI for Point Multiplier Value
    lblPointMultiplier_Value.Caption = PointMultiplier

    'Updates the Distortion Bar with Values
    prbDistortion.Value = DistortionValue
    lblDistortion_Value.Caption = Round(CDbl(DistortionValue / 2), 2)
    
    'Updates the Slow Motion Bar with Values
    prbSlowMotion.Value = SlowMotionValue
    lblSlowMotion_Value.Caption = Round(CDbl(SlowMotionValue / 2), 2)
    
    'Updates the EMP Bar with Values
    prbEMP.Value = EMPValue
    lblEMP_Value.Caption = Round(CDbl(EMPValue / 2), 2)
    
    '--- EMP CODE ---
    
    If isEMPTrigerred = True Then 'Checks if EMP has been triggered
    
        'These Function Calls make the Enemy Lose track of Player
        Call EnemyUserLoseTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
        Call EnemyUserLoseTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
        Call EnemyUserLoseTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
        Call EnemyUserLoseTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
        Call EnemyUserLoseTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
        Call EnemyUserLoseTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
            
        EMPValue = 0 'Empties Bar
        timEnemyMotion.Interval = 200 'Slows all Enemies Down
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                    + vbNewLine _
                    + vbNewLine _
                    + ">>> EMP has been activated. Enemies have slowed down for 5 Seconds! The Enemy's Tracking system has been scrambled!"
        
        lblEMP_OffOn.Caption = "Used" 'Updates UI
        
        isEMPTrigerred = False 'Sets the EMP Trigerred to False
        timEMP.Enabled = True 'Enables Timer
    
    End If
    
    '--- SLOW MOTION CODE ---
    
    'These Statements check the state of Slow Motion (on or off) and then do the relevant things
    If isSlowMotion = True Then
    
        If isEMPOn = False Then
    
            If SlowMotionValue > 0 Then 'Safety check
            
                timEnemyMotion.Interval = 40 'Slows Enemies
                timUserMotion.Interval = 25 'Slows User (yet still twice as fast as enemies)
                lblSlowMotion_OffOn.Caption = "ON" 'Updates UI
                SlowMotionValue = SlowMotionValue - 0.5 'Depletes Bar
            
            ElseIf SlowMotionValue <= 0 Then 'If Less than Zero then
                
                timEnemyMotion.Interval = 1 'Returns Normality
                timUserMotion.Interval = 1 'Returns Normality
                lblSlowMotion_OffOn = "OUT" 'Updates UI
                isSlowMotion = False 'Exitx Decision Statement
            
            End If
            
        End If
        
    Else
    
        If SlowMotionValue = 0 Then 'Custom check
        
            If isEMPOn = False Then
            
                timEnemyMotion.Interval = 1 'Returns Normality
                timUserMotion.Interval = 1 'Returns Normality
                lblSlowMotion_OffOn.Caption = "OUT" 'Updates UI
            
            End If
            
        Else 'Else
        
            If isEMPOn = False Then
            
                timEnemyMotion.Interval = 1 'Returns Normality
                timUserMotion.Interval = 1 'Returns Normality
                lblSlowMotion_OffOn.Caption = "OFF" 'Updates UI
            
            End If
            
        End If 'End Decision Making
        
    End If 'Ends Slow Motion Decision Making
    
    '--- POWER UP CODE ---
    
    'This checks if The PowerUp is Visible _
    and then checks if User Collides with it
    If PowerUpBit.Visible = True Then

        If (PowerUpBitNumber) = 2 Or (PowerUpBitNumber) = 30 Then 'Checks ID of Power Up (SLOW MO FILL UP)

            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
                
                'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PowerUpPickUp = True 'Sets this to True
                
                SlowMotionValue = 200 'Fills Up Slow Motion Bar
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Your Slow Motion Bar has been Refilled!"
                
            End If
            
        ElseIf (PowerUpBitNumber) = 6 Or (PowerUpBitNumber) = 27 Then 'Checks ID of Power Up (DISTORTION FILL UP)

            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
                
                'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PowerUpPickUp = True 'Sets this to True
                
                DistortionValue = 200  'Fills Up Distortion Bar
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Your Distortion Bar has been Refilled!"
                
            End If
            
        ElseIf (PowerUpBitNumber) = 10 Or (PowerUpBitNumber) = 23 Then 'Checks ID of Power Up (ENEMY SPEED UP BY 10%)
                
            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
            
                'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PowerUpPickUp = True 'Sets this to True
                
                Call EnemyIncreaseSpeed(EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyIncreaseSpeed(EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyIncreaseSpeed(EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyIncreaseSpeed(EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyIncreaseSpeed(EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyIncreaseSpeed(EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Enemy Bit Speeds have increased by 10%!"
                            
            End If
            
        ElseIf (PowerUpBitNumber) = 18 Or (PowerUpBitNumber) = 15 Then 'Checks ID of Power Up (ENEMY SPEED DOWN BY 10%)
                
            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
            
                'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PowerUpPickUp = True 'Sets this to True
                
                Call EnemyDecreaseSpeed(EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyDecreaseSpeed(EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyDecreaseSpeed(EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyDecreaseSpeed(EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyDecreaseSpeed(EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyDecreaseSpeed(EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Enemy Bit Speeds have decreased by 10%!"
                            
            End If
            
        ElseIf (PowerUpBitNumber) = 14 Or (PowerUpBitNumber) = 19 Then 'Checks ID of Power Up (LEVEL MODIFIER)

            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
                
                'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
                Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
                PowerUpPickUp = True 'Sets this to True
                
                timPowerUpResultant = False  'Stops Timer
                timPowerUpResultant = True 'Then Restarts It
                timPowerUp = False 'Stop Power Ups during this period
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                
                isWall = True 'Sets to True
                
                'Random Number from 1 to 4
                WallNumber = CInt(Int((4 - 1 + 1) * Rnd() + 1))
                
                'These Call the Movement in wall, and then changes positions if necessary
                Call WallModifier(WallNumber)
                Call WallBitPosition(WallNumber, 6, 0)
                
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> The Level has been Modified for 5 Seconds..."
                
            End If
            
        ElseIf (PowerUpBitNumber) = 3 Or (PowerUpBitNumber) = 29 Then 'Checks ID of Power Up (ENEMY FREEZE)

            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
                
                PowerUpPickUp = True 'Sets this to True
                
                timPowerUpResultant = False  'Stops Timer
                timPowerUpResultant = True 'Then Restarts It
                timPowerUp = False 'Stop Power Ups during this period
                
                timEnemyMotion.Enabled = False 'Stops Enemys
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up Spawn
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> The Enemy has been frozen for 5 Seconds..."
                
            End If
            
        ElseIf (PowerUpBitNumber) = 11 Or (PowerUpBitNumber) = 22 Then 'Checks ID of Power Up (PLAYER FREEZE)

            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
                
                PowerUpPickUp = True 'Sets this to True
                
                timPowerUpResultant = False  'Stops Timer
                timPowerUpResultant = True 'Then Restarts It
                timPowerUp = False 'Stop Power Ups during this period
                
                timUserMotion.Enabled = False 'Stops Enemys
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up Spawn
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You have been frozen for 5 Seconds..."
                
            End If
            
        ElseIf (PowerUpBitNumber) = 26 Or (PowerUpBitNumber) = 7 Then 'Checks ID of Power Up (PLAYER TELEPORT)

            If Abs(UserBitX - PowerUpBitX) < 300 And Abs(UserBitY - PowerUpBitY) < 300 Then
                
                PowerUpPickUp = True 'Sets this to True
                
                timPowerUpResultant = False  'Stops Timer
                timPowerUpResultant = True 'Then Restarts It
                timPowerUp = False 'Stop Power Ups during this period
                
                isPlayerTP = True  'Sets Teleportation to True
                
                PointMultiplier = PointMultiplier + 1 'Increments Point Multiplier
                TotalPoints = TotalPoints + 50 'Adds a Total of 50 Points on Score
                timPowerUp.Interval = 1000 'Sets Back the Original Speed of Power Up Spawn
                PowerUpBit.Visible = False 'Sets the Power Up Bit to false
                
                'Text on Log to inform Player
                txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You will get Teleported to a random position in 5 Seconds..."
                
            End If
                
        End If
        
    End If 'Ends Decision Process
    
    'This Checks if User Collected Power up _
    and if it didn't then it sets the point multiplier _
    back to 1
    If PowerUpBit.Visible = False Then
    
        If PowerUpPickUp = False Then 'If Not Collected
        
            PointMultiplier = 1 'Sets it Back to 1
            
        End If
    
    End If 'End Decision Process
    
    '--- DISTORTION CODE ---
    
    'These Statements check the state of Distortion (on or off) and then do the relevant things
    If isDistortion = False Then
        
        UserBit.Visible = True
        DistortionBit.Visible = False 'Doens't Show Distortion
        
        If DistortionValue > 0 Then 'Check if Distortion is still there
        
            lblDistortion_OffOn = "OFF" 'Update UI
        
        Else
        
            lblDistortion_OffOn = "OUT" 'Update UI
            
        End If
        
    ElseIf isDistortion = True Then 'If ON then
    
        DistortionBit.Visible = True 'Show Distortion
        lblDistortion_OffOn = "ON" 'Update UI
        
        UserBit.Visible = False 'Set the User Bit to be False
        
        DistortionValue = DistortionValue - 0.5 'Reduce Distortion Value
        
        If DistortionValue < 1.5 Then 'Safety Check to avoid Error
            
            DistortionBit.Visible = False 'Don't show Distortion
            isDistortion = False 'Distortion is Off
            DistortionValue = 0 'Noting left in Bar
                
        End If

    End If 'End Selection Process
    
    If isEnemyDistortion = True Then 'Checks if Enemy Collision w/ Player while Distortion
    
        timEnemyMotion.Interval = 40 'Slows Enemy
        
    ElseIf isEnemyDistortion = False Then 'if Not
    
        If isSlowMotion = False Then 'Checks if Slow Motion is on
            
            If isEMPOn = False Then
    
                timEnemyMotion.Interval = 1 'Returns Normality
            
            End If 'End Choices
            
        End If 'End Choices
    
    End If 'End Selection

End Sub
Private Sub timPowerUp_Timer()

    'These Create Random (X,Y) Positions for the Power Up
    Call PowerUpBitMovement
    
    'PowerUpBit.Caption = PowerUpBitNumber 'A Way to display the Power Up ID when Testing
    
    If (PowerUpBitNumber) = 2 Or (PowerUpBitNumber) = 30 Then 'If Power Up ID is 2 or 30 (SLOW MOTION FILL UP)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 10 Or (PowerUpBitNumber) = 23 Then 'If Power Up ID is 10 or 23 (ENEMY SPEED UP BY 10%)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 18 Or (PowerUpBitNumber) = 15 Then 'If Power Up ID is 18 or 15 (ENEMY SPEED DOWN BY 10%)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 14 Or (PowerUpBitNumber) = 19 Then 'If Power Up ID is 14 or 19 (LEVEL MODIFIER)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 3 Or (PowerUpBitNumber) = 29 Then 'If Power Up ID is 3 or 29 (ENEMY FREEZE)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 6 Or (PowerUpBitNumber) = 27 Then 'If Power Up ID is 6 or 27 (DISTORTION BAR FILL)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 11 Or (PowerUpBitNumber) = 22 Then 'If Power Up ID is 11 or 22 (PLAYER FREEZE)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    ElseIf (PowerUpBitNumber) = 26 Or (PowerUpBitNumber) = 7 Then 'If Power Up ID is 26 or 7 (PLAYER TELEPORT)
    
        timPowerUp.Interval = 5000 'Changes time
        PowerUpPickUp = False 'Sets the Pick Up to False, just before visibility change
        PowerUpBit.Visible = True 'Sets the Visibilty to True
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> A Power Up has Appeared in the Arena!"
                            
    Else 'If No Match for ID
        
        PowerUpBit.Visible = False 'If Power Up ID is no Match, visible is false
        
    End If 'Ends selection of ID's

End Sub

Private Sub timPowerUpResultant_Timer()

    If isWall = True Then 'if there is a Wall
    
        isWall = False 'Sets to False
        
        'Wall (X, Y) Values
        W(1) = 125
        W(2) = 125
        W(3) = 125
        W(4) = 6275
        W(5) = 125
        W(6) = 125
        W(7) = 13700
        W(8) = 125
        
        timPowerUp = True 'Re-Enables Power Ups
        frmSurvivalEndless.Refresh 'Refreshes Form
        
    End If
    
    If timEnemyMotion.Enabled = False Then  'Checks if Enemy is Frozen
    
        timEnemyMotion.Enabled = True 'Then sets their motion back moving
        timPowerUp = True 'Re Enables Power Ups
        
        'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
        
    End If
    
    If timUserMotion.Enabled = False Then  'Checks if Player is Frozen
    
        timUserMotion.Enabled = True 'Then sets their motion back moving
        timPowerUp = True 'Re Enables Power Ups
        
        'Call Sub-Routine to calculate new path for Enemy Bits towards User Bit
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit1X, EnemyBit1Y, EnemyBit1XSpeed, EnemyBit1YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit2X, EnemyBit2Y, EnemyBit2XSpeed, EnemyBit2YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit3X, EnemyBit3Y, EnemyBit3XSpeed, EnemyBit3YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit4X, EnemyBit4Y, EnemyBit4XSpeed, EnemyBit4YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit5X, EnemyBit5Y, EnemyBit5XSpeed, EnemyBit5YSpeed)
        Call EnemyUserTarget(UserBitX, UserBitY, EnemyBit6X, EnemyBit6Y, EnemyBit6XSpeed, EnemyBit6YSpeed)
                
    End If
    
    If isPlayerTP = True Then 'Checks if Player Teleportation has been enabled
    
        'Recalculates Players Co-Ordinates...
        UserBitX = CInt(Int((12600 - 300 + 1) * Rnd() + 300))
        UserBitY = CInt(Int((5500 - 300 + 1) * Rnd() + 300))
        
        isPlayerTP = False 'Set back to false
        
        timPowerUp = True 'Re Enables Power Ups
        
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                    + vbNewLine _
                    + vbNewLine _
                    + ">>> Teleported!"
    
    End If

End Sub

Private Sub timRegeneration_Timer() 'New Timer to Naturally Regenerate Bars and Check Walls.
    
    If SlowMotionValue <= 195 Then 'Checks if Less
    
        SlowMotionValue = SlowMotionValue + 5 'Adds More!
        
    End If 'Ends Choices
    
    If DistortionValue <= 195 Then 'Checks if Less
    
        DistortionValue = DistortionValue + 5 'Adds More!
        
    End If 'Ends Choices

End Sub

Private Sub timSurvivalEndless_Timer()
    
    lblScore.Caption = TotalPoints 'Sets Score in UI
    
    If (TotalPoints > 30999) Then 'Checks if Player has achieved the necessary 'win' conditions (To stop Overflow).
        
        isExploded = True
    
        isEnemyHit = False 'Set it back to false
        
        timUserMotion.Enabled = False
        timEnemyMotion.Enabled = False
        timGameMechanics.Enabled = False
        timEMP.Enabled = False
        timPowerUp.Enabled = False
        timRegeneration.Enabled = False
        timPowerUpResultant.Enabled = False
        timSurvivalEndless.Enabled = False
        timWarpMineBit.Enabled = False
        
        UserBit.Visible = False
        DistortionBit.Visible = False
        EnemyBit1.Visible = False
        EnemyBit2.Visible = False
        EnemyBit3.Visible = False
        EnemyBit4.Visible = False
        EnemyBit5.Visible = False
        EnemyBit6.Visible = False
        PowerUpBit.Visible = False
        WarpMineBit1.Visible = False
        WarpMineBit2.Visible = False
        Explosion1.Visible = True
        
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Self-Destruction... DONE"
                            
        fraRestart.Visible = True
        lblGameOver.Caption = "You have beaten Survival: Endless! You have finished all the Survival Levels, but have you done ZomBits?"
        
        If Explosion1.Left > 7000 Then
        
            fraRestart.Left = 240
            cmdReset.Caption = "Go to ZomBits!"
            cmdReset.SetFocus
            
        Else
        
            fraRestart.Left = 8650
            cmdReset.Caption = "Go to ZomBits!"
            cmdReset.SetFocus
        
        End If
        
    End If
        
    Call checkEnemyUserCollision 'Checks Collisions between Enemy and Player
    
    If (isEnemyHit) = True And (isDistortion) = True Then 'If Collision yet Distorted then
        
        isEnemyDistortion = True 'Set this to True
        isEnemyHit = False 'Set it back to false
        
    ElseIf (isEnemyHit) = True And (isDistortion) = False Then 'If Collision then
    
        isExploded = True
    
        isEnemyHit = False 'Set it back to false
        
        timUserMotion.Enabled = False
        timEnemyMotion.Enabled = False
        timGameMechanics.Enabled = False
        timEMP.Enabled = False
        timPowerUpResultant.Enabled = False
        timRegeneration.Enabled = False
        timPowerUp.Enabled = False
        timSurvivalEndless.Enabled = False
        timWarpMineBit.Enabled = False
        
        UserBit.Visible = False
        DistortionBit.Visible = False
        EnemyBit1.Visible = False
        EnemyBit2.Visible = False
        EnemyBit3.Visible = False
        EnemyBit4.Visible = False
        EnemyBit5.Visible = False
        EnemyBit6.Visible = False
        PowerUpBit.Visible = False
        WarpMineBit1.Visible = False
        WarpMineBit2.Visible = False
        Explosion1.Visible = True
        
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> BOOM!"
                            
        fraRestart.Visible = True
        lblGameOver.Caption = "Game Over. Your Bit has been Obliterated because you collided with your Enemy. You have lost the battle."
        
        If Explosion1.Left > 7000 Then
        
            fraRestart.Left = 240
            cmdReset.SetFocus
            
        Else
        
            fraRestart.Left = 8650
            cmdReset.SetFocus
        
        End If
        
    ElseIf (isBurntOut) = True Then
    
        isBurntOut = False 'Set it back to false
        
        isExploded = True
    
        isEnemyHit = False 'Set it back to false
        
        timUserMotion.Enabled = False
        timEnemyMotion.Enabled = False
        timGameMechanics.Enabled = False
        timEMP.Enabled = False
        timPowerUpResultant.Enabled = False
        timRegeneration.Enabled = False
        timPowerUp.Enabled = False
        timSurvivalEndless.Enabled = False
        timWarpMineBit.Enabled = False
        
        UserBit.Visible = False
        DistortionBit.Visible = False
        EnemyBit1.Visible = False
        EnemyBit2.Visible = False
        EnemyBit3.Visible = False
        EnemyBit4.Visible = False
        EnemyBit5.Visible = False
        EnemyBit6.Visible = False
        PowerUpBit.Visible = False
        WarpMineBit1.Visible = False
        WarpMineBit2.Visible = False
        Explosion1.Visible = True
        
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> BOOM!"
                            
        fraRestart.Visible = True
        lblGameOver.Caption = "Game Over. Your Bit has been Incinerated because your engines became too hot (Your Speed was more than 180) . You have lost the battle."
        
        If Explosion1.Left > 7000 Then
        
            fraRestart.Left = 240
            cmdReset.SetFocus
            
        Else
        
            fraRestart.Left = 8650
            cmdReset.SetFocus
        
        End If
    
    Else
        
        isEnemyDistortion = False 'if Nothing, False
        
    End If
    
End Sub

Private Sub timWarpMineBit_Timer()
    
    If TotalPoints > 400 Then
         
        If EndlessSurvivalProgression = 3 Then
        
            EndlessSurvivalProgression = 4
        
        End If
            
        WarpMineBit1.Visible = True
    
        'This Function Calls set Random Position
        Call WarpMineBitPosition(WarpMineBit1X, WarpMineBit1Y)
        
        'This Call checks if the Warp Bit is actually too near the player
        Call checkWarpBitSpawnProximity
        
        If isWarpMineProximal = True Then
        
            'Calls Random PLace Again
            Call WarpMineBitPosition(WarpMineBit1X, WarpMineBit1Y)
            isWarpMineProximal = False
            
        Else
        
            'Assigns Position to Bits
            WarpMineBit1.Left = WarpMineBit1X
            WarpMineBit1.Top = WarpMineBit1Y
        
        End If
    End If
    If TotalPoints > 800 Then
            
            If EndlessSurvivalProgression = 5 Then
        
                EndlessSurvivalProgression = 6
        
            End If
            WarpMineBit2.Visible = True
        
            'This Function Calls set Random Position
            Call WarpMineBitPosition(WarpMineBit2X, WarpMineBit2Y)
            
            'This Call checks if the Warp Bit is actually too near the player
            Call checkWarpBitSpawnProximity
            
            If isWarpMineProximal = True Then
            
                'Calls Random PLace Again
                Call WarpMineBitPosition(WarpMineBit2X, WarpMineBit2Y)
                isWarpMineProximal = False
                
            Else
            
                'Assigns Position to Bits
                WarpMineBit2.Left = WarpMineBit2X
                WarpMineBit2.Top = WarpMineBit2Y
            
            End If
            
    End If
        
    

End Sub

'Skips to bottom text of System Log
Private Sub txtSystemLog_Change()
    txtSystemLog.SelStart = Len(txtSystemLog.Text)
    
        txtSystemLog.SelLength = 0
End Sub

Private Sub timUserMotion_Timer()
    
    'Allocate the Position of User Bit to Stored variables
    UserBit.Left = UserBitX
    UserBit.Top = UserBitY
    Explosion1.Top = UserBit.Top - 350
    Explosion1.Left = UserBit.Left - 350
    
    TopFire1.Top = UserBit.Top - 300
    TopFire1.Left = UserBit.Left
    LeftFire1.Top = UserBit.Top
    LeftFire1.Left = UserBit.Left - 300
    RightFire1.Top = UserBit.Top
    RightFire1.Left = UserBit.Left + 300
    DownFire1.Top = UserBit.Top + 300
    DownFire1.Left = UserBit.Left
    
    If TopFire = True Then
    
        TopFire1.Visible = True
        TopFire = False
        
    Else
    
        TopFire1.Visible = False
        
    End If
    
    If RightFire = True Then
    
        RightFire1.Visible = True
        RightFire = False
        
    Else
    
        RightFire1.Visible = False
        
    End If
    If LeftFire = True Then
    
        LeftFire1.Visible = True
        LeftFire = False
        
    Else
    
        LeftFire1.Visible = False
        
    End If
    If DownFire = True Then
    
        DownFire1.Visible = True
        DownFire = False
        
    Else
    
        DownFire1.Visible = False
        
    End If
    'Allocate the Position of Shield to stored Variables
    DistortionBit.Top = DistortionBitY
    DistortionBit.Left = DistortionBitX
    
    txtUpSpeed.Text = UserBitUpSpeed
    txtDownSpeed.Text = UserBitDownSpeed
    txtLeftSpeed.Text = UserBitLeftSpeed
    txtRightSpeed.Text = UserBitRightSpeed
    
    'Call Sub-Routine to increase/decrease speed based on keys, hence movement
    Call UserBitMovement(UserBitX, UserBitY, UserBitXSpeed, UserBitYSpeed)
    
    'Call Sub-Routine to map Shields to User Bit
    Call DistortionBitMovement(UserBitX, UserBitY, DistortionBitX, DistortionBitY)
    
    'Call Sub-Routine to check Wall Collisons
    Call checkUserWallCollison
    
    'Call Sub-Routine to check user Speeds
    Call CheckUserSpeeds

End Sub





