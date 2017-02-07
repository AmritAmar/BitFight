VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitFight! - Main Menu"
   ClientHeight    =   8376
   ClientLeft      =   912
   ClientTop       =   1584
   ClientWidth     =   17028
   Icon            =   "BitFight! - Main Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BitFight! - Main Menu.frx":058A
   ScaleHeight     =   8376
   ScaleWidth      =   17028
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About Game"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   1200
      TabIndex        =   38
      Top             =   600
      Width           =   2172
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   1200
      MaskColor       =   &H00E0E0E0&
      Picture         =   "BitFight! - Main Menu.frx":7A0EC
      TabIndex        =   1
      Top             =   2640
      Width           =   2172
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   1200
      TabIndex        =   37
      Top             =   4680
      Width           =   2172
   End
   Begin VB.Timer timBG 
      Interval        =   10
      Left            =   1800
      Top             =   120
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play!"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   10560
      TabIndex        =   36
      Top             =   2640
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton cmdSEndless 
      Caption         =   "Endless"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   7320
      TabIndex        =   35
      Top             =   5160
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton cmdSHard 
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   7320
      TabIndex        =   34
      Top             =   3600
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton cmdSMedium 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   7320
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton cmdSEasy 
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   7320
      TabIndex        =   32
      Top             =   480
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton cmdZomBits 
      Caption         =   "ZomBits"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4200
      TabIndex        =   31
      Top             =   3600
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton cmdSurvival 
      Caption         =   "Survival"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4200
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Timer timRegeneration 
      Interval        =   5000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer timGameMechanics 
      Interval        =   1
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer timMainMenu 
      Interval        =   1
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer timEnemyMotion 
      Interval        =   1
      Left            =   360
      Top             =   120
   End
   Begin VB.Frame frmHolderLog 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6372
      Left            =   13800
      TabIndex        =   6
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
         TabIndex        =   7
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
            TabIndex        =   8
            Top             =   360
            Width           =   2892
         End
      End
   End
   Begin VB.Frame frmHolderInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   2052
      Left            =   120
      TabIndex        =   5
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
         TabIndex        =   9
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
            TabIndex        =   15
            Text            =   "x"
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
            TabIndex        =   14
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
            TabIndex        =   13
            Text            =   "x"
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
            TabIndex        =   12
            Text            =   "x"
            Top             =   1080
            Width           =   1092
         End
         Begin MSComctlLib.ProgressBar prbDistortion 
            Height          =   372
            Left            =   6480
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   16
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
            TabIndex        =   29
            Top             =   360
            Width           =   972
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   960
            Width           =   2652
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
            TabIndex        =   20
            Top             =   960
            Width           =   492
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   1320
            Width           =   1092
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
            TabIndex        =   17
            Top             =   1320
            Width           =   972
         End
      End
   End
   Begin VB.Timer timUserMotion 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF0000&
      X1              =   2280
      X2              =   2280
      Y1              =   1440
      Y2              =   2640
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      X1              =   2280
      X2              =   2280
      Y1              =   3480
      Y2              =   4680
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      Visible         =   0   'False
      X1              =   6360
      X2              =   10560
      Y1              =   4080
      Y2              =   3120
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      Visible         =   0   'False
      X1              =   9480
      X2              =   10560
      Y1              =   5640
      Y2              =   3120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      Visible         =   0   'False
      X1              =   9480
      X2              =   10560
      Y1              =   4080
      Y2              =   3120
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      Visible         =   0   'False
      X1              =   9480
      X2              =   10560
      Y1              =   2520
      Y2              =   3120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      Visible         =   0   'False
      X1              =   9480
      X2              =   10560
      Y1              =   840
      Y2              =   3120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      Visible         =   0   'False
      X1              =   7320
      X2              =   6360
      Y1              =   5640
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      Visible         =   0   'False
      X1              =   7320
      X2              =   6360
      Y1              =   4080
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      Visible         =   0   'False
      X1              =   7320
      X2              =   6360
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      Visible         =   0   'False
      X1              =   6360
      X2              =   7320
      Y1              =   2280
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   4200
      X2              =   3360
      Y1              =   4080
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   3360
      X2              =   4200
      Y1              =   3120
      Y2              =   2280
   End
   Begin VB.Label BottomCounterMeasure 
      Height          =   96
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   96
   End
   Begin VB.Label RightCounterMeasure 
      Height          =   96
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   96
   End
   Begin VB.Label LeftCounterMeasure 
      Height          =   96
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   96
   End
   Begin VB.Label TopCounterMeasure 
      BackColor       =   &H00FFFFFF&
      Height          =   96
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   96
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R As Integer
Dim B As Integer
Dim G As Integer

Private Sub cmdAbout_Click()

    Unload Me
    frmAbout.Show

End Sub

Private Sub cmdInstructions_Click()

    Unload Me
    frmInstructions.Show

End Sub

Private Sub cmdPlay_Click()

    If MainMenuProgression = 1 Then

        Unload Me
        frmSurvivalEasy.Show
        
    ElseIf MainMenuProgression = 2 Then

        Unload Me
        frmSurvivalMedium.Show
        
    ElseIf MainMenuProgression = 3 Then

        Unload Me
        frmSurvivalHard.Show
        
    ElseIf MainMenuProgression = 4 Then

        Unload Me
        frmSurvivalEndless.Show
        
    ElseIf MainMenuProgression = 5 Then

        Unload Me
        frmZomBits.Show
        
    End If

End Sub

Private Sub cmdSEasy_Click()

        txtSystemLog.Text = txtSystemLog.Text + _
                            vbNewLine _
                            + vbNewLine _
                            + ">>> You have selected Survival: Easy!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You will have 2 Enemies against you. You have to survive until 1200 Points are reached and you have a 4 Point Multiplier!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Click 'Play!' when Ready!" _

    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = True
    Line8.Visible = False
    Line9.Visible = False
    Line10.Visible = False
    Line11.Visible = False
    
    cmdSEasy.Visible = True
    cmdSMedium.Visible = True
    cmdSHard.Visible = True
    cmdSEndless.Visible = True
    cmdPlay.Visible = True
    
    cmdPlay.SetFocus
    
    MainMenuProgression = 1

End Sub

Private Sub cmdSEndless_Click()
    
    txtSystemLog.Text = txtSystemLog.Text + _
                            vbNewLine _
                            + vbNewLine _
                            + ">>> You have selected Survival: Endless!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You will be put in an arena, where the number of enemies increases as the level progresses. You have to survive..." _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Click 'Play!' when Ready!"
                                
    
    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Line10.Visible = True
    Line11.Visible = False
    
    cmdSEasy.Visible = True
    cmdSMedium.Visible = True
    cmdSHard.Visible = True
    cmdSEndless.Visible = True
    cmdPlay.Visible = True
    
    cmdPlay.SetFocus

    MainMenuProgression = 4

End Sub

Private Sub cmdSHard_Click()

        txtSystemLog.Text = txtSystemLog.Text + _
                            vbNewLine _
                            + vbNewLine _
                            + ">>> You have selected Survival: Hard!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You will have 4 Enemies against you and 1 Warp Mine. You have to survive until 2000 Points are reached and you have a 12 Point Multiplier!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Click 'Play!' when Ready!"


    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = True
    Line10.Visible = False
    Line11.Visible = False
    
    cmdSEasy.Visible = True
    cmdSMedium.Visible = True
    cmdSHard.Visible = True
    cmdSEndless.Visible = True
    cmdPlay.Visible = True
    
    cmdPlay.SetFocus
    
    MainMenuProgression = 3

End Sub

Private Sub cmdSMedium_Click()

        txtSystemLog.Text = txtSystemLog.Text + _
                            vbNewLine _
                            + vbNewLine _
                            + ">>> You have selected Survival: Medium!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> You will have 3 Enemies against you and 1 Warp Mine. You have to survive until 1600 Points are reached and you have a 8 Point Multiplier!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Click 'Play!' when Ready!" _


    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = False
    Line8.Visible = True
    Line9.Visible = False
    Line10.Visible = False
    Line11.Visible = False
    
    cmdSEasy.Visible = True
    cmdSMedium.Visible = True
    cmdSHard.Visible = True
    cmdSEndless.Visible = True
    cmdPlay.Visible = True
    
    cmdPlay.SetFocus
    
    MainMenuProgression = 2

End Sub

Private Sub cmdStart_Click()

    txtSystemLog.Text = txtSystemLog.Text + _
                    vbNewLine _
                    + vbNewLine _
                    + ">>> BitFight!" _
                    + vbNewLine _
                    + vbNewLine _
                    + ">>> This is a Game about Survival. Avoid Enemy Bits and Warp Mine by controlling your Bit Skillfully. Use W,A,S,D or the Arrow Keys to Move. You can also activate Slow Motion by using the Space Bar, Distortion by pressing 'Q' or 'N', and finally EMP by pressing 'E' or 'B'..." _
                    + vbNewLine _
                    + vbNewLine _
                    + ">>> Choose the Game Mode you want to Play!" _

    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = False
    Line4.Visible = False
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Line10.Visible = False
    Line11.Visible = False
    
    cmdSEasy.Visible = False
    cmdSMedium.Visible = False
    cmdSHard.Visible = False
    cmdSEndless.Visible = False
    cmdPlay.Visible = False
    
    cmdSurvival.SetFocus

End Sub

Private Sub cmdSurvival_Click()

    txtSystemLog.Text = txtSystemLog.Text + _
                    vbNewLine _
                    + vbNewLine _
                    + ">>> You have selected Survival!" _
                    + vbNewLine _
                    + vbNewLine _
                    + ">>> This is a Classic Gamemode, where you have to avoid all the enemies in the Arena, Get Power Ups and Survive!" _
                    + vbNewLine _
                    + vbNewLine _
                    + ">>> Choose Your Difficulty!" _


    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Line10.Visible = False
    Line11.Visible = False
    
    cmdSEasy.Visible = True
    cmdSMedium.Visible = True
    cmdSHard.Visible = True
    cmdSEndless.Visible = True
    cmdPlay.Visible = False
    
    cmdSEasy.SetFocus

End Sub

Private Sub cmdZomBits_Click()

    txtSystemLog.Text = txtSystemLog.Text + _
            vbNewLine _
            + vbNewLine _
            + ">>> You have selected ZomBits!" _
            + vbNewLine _
            + vbNewLine _
            + ">>> Zombits is a game about survival, but only from 1 side of the Arena. The Zombits will come Hard and Fast. Goodluck!" _
            + vbNewLine _
            + vbNewLine _
            + ">>> Click 'Play!' when you are Ready!" _


    cmdSurvival.Visible = True
    cmdZomBits.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = False
    Line4.Visible = False
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Line10.Visible = False
    Line11.Visible = True
    
    cmdSEasy.Visible = False
    cmdSMedium.Visible = False
    cmdSHard.Visible = False
    cmdSEndless.Visible = False
    cmdPlay.Visible = True
    
    cmdPlay.SetFocus
    
    MainMenuProgression = 5

End Sub

Private Sub Form_Load()

    Call StartUpMainMenu
    
    'cmdStart.SetFocus
    
    SlowMotionValue = 200
    prbDistortion = 200
    prbEMP = 200
    
    lblDistortion_Value = 100
    lblSlowMotion_Value = 100
    lblEMP_Value = 100
    
    'Change The Text Box to Show the Following Intructions in the Side log
    txtSystemLog.Text = " - LOG INITIALIZED - " _
                            + vbNewLine + vbNewLine _
                            + ">>> BitFight Routine Started!" _
                            + vbNewLine _
                            + ">>> Welcome to BitFight!" _
                            + vbNewLine _
                            + vbNewLine _
                            + ">>> Click on 'Instructions' or Choose your options to play the game!" _
                            
    'Just a String for the Score Box
    lblScore.Caption = "BitFight - Beta v1.0"
    'Call the Start up Routine
    
    'General Things to do
    prbSlowMotion.Value = SlowMotionValue 'Set Slow Mo Value
    
    'Colour Allocation

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
    timMainMenu.Enabled = True
    timGameMechanics.Enabled = True

End Sub

Private Sub Form_Paint()
    
    'Draw all the Walls
    Line (125, 125)-(125 + 13572, 125 + 15), RGB(0, 255, 0), BF
    Line (125, 6275)-(125 + 13572, 6275 + 15), RGB(0, 255, 0), BF
    Line (125, 125)-(125 + 15, 125 + 6156), RGB(0, 255, 0), BF
    Line (13700, 125)-(13700 + 15, 125 + 6156), RGB(0, 255, 0), BF
    
End Sub

'Skips to bottom text of System Log
Private Sub txtSystemLog_Change()
    txtSystemLog.SelStart = Len(txtSystemLog.Text)
    
        txtSystemLog.SelLength = 0
End Sub

