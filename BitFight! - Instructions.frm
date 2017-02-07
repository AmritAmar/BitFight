VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInstructions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitFight! - Instructions"
   ClientHeight    =   8376
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   17016
   Icon            =   "BitFight! - Instructions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BitFight! - Instructions.frx":058A
   ScaleHeight     =   8376
   ScaleWidth      =   17016
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Main Menu"
      Height          =   612
      Left            =   6360
      TabIndex        =   37
      Top             =   2760
      Width           =   2172
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
            TabIndex        =   8
            Text            =   "x"
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
            TabIndex        =   7
            Text            =   "x"
            Top             =   1080
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
            TabIndex        =   6
            Text            =   "x"
            Top             =   1080
            Width           =   1092
         End
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
            TabIndex        =   5
            Text            =   "x"
            Top             =   480
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
         Begin VB.Label lblEMP_OffOn 
            BackColor       =   &H00000000&
            Caption         =   "OFF"
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
            TabIndex        =   24
            Top             =   1320
            Width           =   972
         End
         Begin VB.Label lblEMP_Value 
            BackColor       =   &H00000000&
            Caption         =   "100"
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
            TabIndex        =   23
            Top             =   1320
            Width           =   1092
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
            TabIndex        =   22
            Top             =   1320
            Width           =   852
         End
         Begin VB.Label lblPointMultiplier_Value 
            BackColor       =   &H00000000&
            Caption         =   "##"
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
            TabIndex        =   21
            Top             =   960
            Width           =   492
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
            TabIndex        =   20
            Top             =   960
            Width           =   2652
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
            TabIndex        =   19
            Top             =   1440
            Width           =   2892
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "BitFight! - Beta"
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
            TabIndex        =   17
            Top             =   840
            Width           =   1692
         End
         Begin VB.Label lblDistortion_Value 
            BackColor       =   &H00000000&
            Caption         =   "100"
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
            TabIndex        =   16
            Top             =   840
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
            TabIndex        =   15
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label lblSlowMotion_Value 
            BackColor       =   &H00000000&
            Caption         =   "100"
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
            TabIndex        =   14
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label lblDistortion_OffOn 
            BackColor       =   &H00000000&
            Caption         =   "OFF"
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
            TabIndex        =   13
            Top             =   840
            Width           =   972
         End
         Begin VB.Label lblSlowMotion_OffOn 
            BackColor       =   &H00000000&
            Caption         =   "OFF"
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
            TabIndex        =   12
            Top             =   360
            Width           =   972
         End
      End
   End
   Begin VB.Frame frmHolderLog 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6372
      Left            =   13920
      TabIndex        =   0
      Top             =   120
      Width           =   3012
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
         Width           =   3012
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
            Width           =   2772
         End
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      X1              =   13800
      X2              =   13800
      Y1              =   240
      Y2              =   6240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   13800
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   13800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   6240
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on any of the Texts to learn more about that Object in the Game."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   972
      Left            =   5400
      TabIndex        =   36
      Top             =   1560
      Width           =   3972
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Height          =   1092
      Left            =   9840
      TabIndex        =   35
      Top             =   2520
      Width           =   3732
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Height          =   1212
      Left            =   9840
      TabIndex        =   34
      Top             =   1200
      Width           =   3012
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Height          =   612
      Left            =   240
      TabIndex        =   33
      Top             =   360
      Width           =   5412
   End
   Begin VB.Label lblInfoPowerUps 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Height          =   852
      Left            =   7800
      TabIndex        =   32
      Top             =   5280
      Width           =   5892
   End
   Begin VB.Shape WarpMineBit1 
      BorderColor     =   &H00FF00FF&
      FillColor       =   &H00FF00FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   300
      Left            =   2160
      Shape           =   5  'Rounded Square
      Top             =   1680
      Width           =   300
   End
   Begin VB.Shape PowerUpBit 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   3  'Vertical Line
      Height          =   300
      Left            =   6480
      Shape           =   5  'Rounded Square
      Top             =   4680
      Width           =   300
   End
   Begin VB.Shape EnemyBit1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   300
      Left            =   12480
      Shape           =   5  'Rounded Square
      Top             =   360
      Width           =   300
   End
   Begin VB.Shape UserBit 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      FillColor       =   &H00FF0000&
      FillStyle       =   6  'Cross
      Height          =   300
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   300
   End
   Begin VB.Label lblDistortioninfo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   360
      TabIndex        =   31
      Top             =   4200
      Width           =   4932
   End
   Begin VB.Label lblEMPinfo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   360
      TabIndex        =   30
      Top             =   5160
      Width           =   5412
   End
   Begin VB.Label lblSlowMotionInfo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   240
      TabIndex        =   29
      Top             =   3360
      Width           =   5172
   End
   Begin VB.Label TopCounterMeasure 
      BackColor       =   &H00FFFFFF&
      Height          =   96
      Left            =   360
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   96
   End
   Begin VB.Label LeftCounterMeasure 
      Height          =   96
      Left            =   240
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   96
   End
   Begin VB.Label RightCounterMeasure 
      Height          =   96
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   96
   End
   Begin VB.Label BottomCounterMeasure 
      Height          =   96
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   96
   End
   Begin VB.Image imgInstructions 
      Height          =   8412
      Left            =   0
      Picture         =   "BitFight! - Instructions.frx":7A0EC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17052
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMainMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub


Private Sub gamemechanics_Timer()

End Sub

Private Sub Label1_Click()
        'Text on Log to inform Player
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> This is an Enemy Bit." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Enemy Bits move in a random direction from the start of the game. They bounce off walls (giving you 1 point each time) and are extremely lethal. Avoid them at all costs." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Enemy Bits can be avoided by using Distortion, as it stops them from attacking you; they pass freely over you. They are also affected greatly by Slow Motion. The most lethal attack is the EMP, which completely disrupts them for a couple of seconds."
    
End Sub

Private Sub Label2_Click()
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> This is a Warp Mine." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> These Mines warp around the screen, appearing in random loacations. They are extremely lethal as colliding with them is fatal. Avoid them at all costs." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Warp Mines do not affect the player if Distortion is used... They warp around the screen every 2 seconds."
        
End Sub

Private Sub Label3_Click()
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> This is your Bit, the User Bit." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You control this Bit in the Game. You have to use WASD or the Arrow Keys to move it. You have to Repeatedly 'Tap' the keys to make a difference, because each tap speeds the Bit up by either 5 Km/h (in Easy and Medium) or 10 Km/h (in Hard, Endless and ZomBits) in the direction it is going." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You have to make sure that this Bit avoids all the Enemy Bits and the Warp Mine Bits. In addition to these fears, you have to make sure your Bit does not exceed 180 Km/H in any direction. The Right Corner of the Screen shows your current speed at any time."

End Sub

Private Sub lblDistortioninfo_Click()
             txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Distortion" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Distortion is a cloaking device, that allows all lethal objects to freely pass over you. It is activated by pressing 'Q'." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Distortion is a get-away plan for unexpected encounters. It can be a life-saver in situations. However, keep an eye on the rapidly decreasing bar below if used. It regenerates naturally and a full bar can be obtained from a Power Up."

End Sub

Private Sub lblEMPinfo_Click()
       txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> EMP (Electro-Magnetic Pulse)" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> The EMP is the most deadly weapon available to you. It is activated by pressing 'E', and your Bit releases an invisible pulse that devastates all Enemy Bits. It slows them down and disrupts their target machanism for a few seconds." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> EMP's are a good way of completely saving your Bit from enemies. However, it can only be used once, until the bar is completely full. Timing is a key-factor when using EMP's."
    
End Sub

Private Sub lblInfoPowerUps_Click()
        txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> This is a Power Up Bit." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You have to collect these randomly spawning Bits throughout the Game. Getting this Bit results in a 50 Point Bonus, an addition to your Point Multiplier and 1 random effect on you, the enemies or the level." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> In the game, it is imperative to collect these to finish the game, as they form part of the objective of getting a certain Point Multiplier. However, beware, because some of the effects are not exactly... Friendly."

End Sub

Private Sub LblMM_Click()
    
    Unload Me
    Call StartUpMainMenu
    frmMainMenu.Show
    
End Sub


Private Sub Form_Load()

    prbDistortion.Value = 200
    prbSlowMotion.Value = 200
    prbEMP.Value = 200

    'Colour Allocation
    UserBit.FillColor = UserBitBGColour
    EnemyBit1.FillColor = EnemyBitBGColour
    WarpMineBit1.FillColor = WarpMineBGColour
    PowerUpBit.FillColor = PowerUpBGColour
    UserBit.BackColor = UserBitBGColour
    EnemyBit1.BackColor = EnemyBitBGColour
    WarpMineBit1.BackColor = WarpMineBGColour
    PowerUpBit.BackColor = PowerUpBGColour
    UserBit.BorderColor = UserBitBGColour
    EnemyBit1.BorderColor = EnemyBitBGColour
    WarpMineBit1.BorderColor = WarpMineBGColour
    PowerUpBit.BorderColor = PowerUpBGColour
    
    If isSolidColours = True Then
    
        UserBit.BackStyle = 1
        EnemyBit1.BackStyle = 1
        WarpMineBit1.BackStyle = 1
        PowerUpBit.BackStyle = 1
        
    End If
    
    If isCircularPowerUp = True Then
    
        PowerUpBit.Shape = ShapeConstants.vbShapeCircle
        
    End If
    
    If isCircularWarpMine = True Then
    
        WarpMineBit1.Shape = ShapeConstants.vbShapeCircle
        
    End If
    
    'Change The Text Box to Show the Following Intructions in the Side log
    txtSystemLog.Text = " - LOG INITIALIZED - " _
                            + vbNewLine + vbNewLine _
                            + ">>> BitFight Routine Started!" _
                            + vbNewLine _
                            + ">>> Welcome to BitFight!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Click the Game Mechanics you want to learn about! Or move your Bit to the Main Menu to Play!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> This is a game about Survival. The Basic Objective is to avoid Enemies and collect Power Ups.. Sounds Easy, doesn't it?"
                            
    
End Sub

Private Sub Form_Paint()

    'Draw all the Walls
    Line (W(1), W(2))-(W(1) + 13572, W(2) + 15), RGB(0, 255, 0), BF
    Line (W(3), W(4))-(W(3) + 13572, W(4) + 15), RGB(0, 255, 0), BF
    Line (W(5), W(6))-(W(5) + 15, W(6) + 6156), RGB(0, 255, 0), BF
    Line (W(7), W(8))-(W(7) + 15, W(8) + 6156), RGB(0, 255, 0), BF
    
End Sub

Private Sub timMainMenu_Timer()

End Sub

Private Sub lblSlowMotionInfo_Click()
         txtSystemLog.Text = txtSystemLog.Text _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Slow Motion" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Slow Motion is a natural power up available to you. It is activated by pressing the 'Space bar'. This slows down all moving objects in the game (You and Enemy Bits). Warp Mines and Power Ups are not affected by this." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Slow Motion is a useful way to escape enemies if your tapping is fast enough. It can be a life-saver in situations. However, keep an eye on the rapidly decreasing bar below if used. It regenerates naturally and a full bar can be obtained from a Power Up."
    
End Sub



Private Sub txtSystemLog_Change()
    
    txtSystemLog.SelStart = Len(txtSystemLog.Text)
    
    txtSystemLog.SelLength = 0
    
End Sub
