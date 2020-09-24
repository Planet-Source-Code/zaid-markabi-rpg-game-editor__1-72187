VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get-Wrong \ Level Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Story"
      Height          =   1335
      Left            =   7680
      TabIndex        =   31
      Top             =   6600
      Width           =   975
      Begin VB.CommandButton Command11 
         Caption         =   "C"
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Music"
      Height          =   855
      Left            =   8760
      TabIndex        =   25
      Top             =   7080
      Width           =   3135
      Begin VB.CommandButton Command9 
         Caption         =   "S"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "P"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":030A
         Left            =   1320
         List            =   "Form1.frx":033B
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Music Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "More Item"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   1575
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Com D"
         Height          =   615
         Index           =   32
         Left            =   120
         Picture         =   "Form1.frx":0372
         Stretch         =   -1  'True
         Tag             =   "DH"
         ToolTipText     =   "E"
         Top             =   240
         Width           =   615
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Com L"
         Height          =   615
         Index           =   31
         Left            =   840
         Picture         =   "Form1.frx":0AD4
         Stretch         =   -1  'True
         Tag             =   "END"
         ToolTipText     =   "E"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map"
      Height          =   1335
      Left            =   1800
      TabIndex        =   10
      Top             =   6600
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "F"
         Height          =   255
         Left            =   5400
         TabIndex        =   30
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "View"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox CRI 
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox PW 
         Height          =   285
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   17
         Text            =   "Enter PassWord"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox LVN 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   495
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Open"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Map"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Cristal Number :"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PassWord Level :"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Level Number :"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Items"
      Height          =   6855
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox PSZ 
         BackColor       =   &H00000000&
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   795
         TabIndex        =   4
         Top             =   5880
         Width           =   855
         Begin VB.Image PS 
            Height          =   855
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.TextBox HLP 
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Button"
         Height          =   495
         Index           =   36
         Left            =   2520
         Picture         =   "Form1.frx":4116
         Stretch         =   -1  'True
         Tag             =   "ICED"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Have Red Key"
         Height          =   495
         Index           =   35
         Left            =   720
         Picture         =   "Form1.frx":5290
         Stretch         =   -1  'True
         Tag             =   "ICER"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Have Blue Key"
         Height          =   495
         Index           =   34
         Left            =   1320
         Picture         =   "Form1.frx":640A
         Stretch         =   -1  'True
         Tag             =   "ICEL"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Have Yellow Key"
         Height          =   495
         Index           =   33
         Left            =   1920
         Picture         =   "Form1.frx":7584
         Stretch         =   -1  'True
         Tag             =   "ICEU"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label HL 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Help For Block"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   2880
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label TOL 
         Caption         =   "Vou"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label TA 
         Caption         =   "Str"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   6240
         Width           =   1815
      End
      Begin VB.Label NAM 
         Caption         =   "Name"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "TEXT :"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Tips Block"
         Height          =   615
         Index           =   30
         Left            =   120
         Picture         =   "Form1.frx":86FE
         Stretch         =   -1  'True
         Tag             =   "TIPS"
         Top             =   5160
         Width           =   615
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Com L"
         Height          =   495
         Index           =   29
         Left            =   2520
         Picture         =   "Form1.frx":93E6
         Stretch         =   -1  'True
         Tag             =   "COM"
         ToolTipText     =   "L"
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Com D"
         Height          =   495
         Index           =   28
         Left            =   1920
         Picture         =   "Form1.frx":AD50
         Stretch         =   -1  'True
         Tag             =   "COM"
         ToolTipText     =   "D"
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Com R"
         Height          =   495
         Index           =   27
         Left            =   1320
         Picture         =   "Form1.frx":C6BA
         Stretch         =   -1  'True
         Tag             =   "COM"
         ToolTipText     =   "R"
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Com U"
         Height          =   495
         Index           =   26
         Left            =   720
         Picture         =   "Form1.frx":E024
         Stretch         =   -1  'True
         Tag             =   "COM"
         ToolTipText     =   "U"
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Boots"
         Height          =   495
         Index           =   25
         Left            =   1320
         Picture         =   "Form1.frx":F98E
         Stretch         =   -1  'True
         Tag             =   "GETICE"
         ToolTipText     =   "E"
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Water Boots"
         Height          =   495
         Index           =   24
         Left            =   720
         Picture         =   "Form1.frx":12C2A
         Stretch         =   -1  'True
         Tag             =   "GETWATER"
         ToolTipText     =   "E"
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Fire Boots"
         Height          =   495
         Index           =   23
         Left            =   120
         Picture         =   "Form1.frx":15EC6
         Stretch         =   -1  'True
         Tag             =   "GETFIRE"
         ToolTipText     =   "E"
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Yellow Key"
         Height          =   495
         Index           =   22
         Left            =   1920
         Picture         =   "Form1.frx":19162
         Stretch         =   -1  'True
         Tag             =   "GETYELLOW"
         ToolTipText     =   "E"
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Blue Key"
         Height          =   495
         Index           =   21
         Left            =   2520
         Picture         =   "Form1.frx":1A7C2
         Stretch         =   -1  'True
         Tag             =   "GETBLUE"
         ToolTipText     =   "E"
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Red Key"
         Height          =   495
         Index           =   20
         Left            =   120
         Picture         =   "Form1.frx":1BE22
         Stretch         =   -1  'True
         Tag             =   "GETRED"
         ToolTipText     =   "E"
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Have Yellow Key"
         Height          =   495
         Index           =   19
         Left            =   1920
         Picture         =   "Form1.frx":1D482
         Stretch         =   -1  'True
         Tag             =   "DOORHY"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Have Blue Key"
         Height          =   495
         Index           =   18
         Left            =   1320
         Picture         =   "Form1.frx":1E0C4
         Stretch         =   -1  'True
         Tag             =   "DOORHB"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Have Red Key"
         Height          =   495
         Index           =   17
         Left            =   720
         Picture         =   "Form1.frx":1ED06
         Stretch         =   -1  'True
         Tag             =   "DOORHR"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Button"
         Height          =   495
         Index           =   16
         Left            =   2520
         Picture         =   "Form1.frx":1F948
         Stretch         =   -1  'True
         Tag             =   "LU"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go LD"
         Height          =   495
         Index           =   15
         Left            =   120
         Picture         =   "Form1.frx":20092
         Stretch         =   -1  'True
         Tag             =   "ICELD"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go LU"
         Height          =   495
         Index           =   14
         Left            =   2520
         Picture         =   "Form1.frx":26B6C
         Stretch         =   -1  'True
         Tag             =   "ICELU"
         ToolTipText     =   "E"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go RD"
         Height          =   495
         Index           =   13
         Left            =   1920
         Picture         =   "Form1.frx":2D642
         Stretch         =   -1  'True
         Tag             =   "ICERD"
         ToolTipText     =   "E"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go RU"
         Height          =   495
         Index           =   12
         Left            =   1320
         Picture         =   "Form1.frx":34120
         Stretch         =   -1  'True
         Tag             =   "ICERU"
         ToolTipText     =   "E"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go DL"
         Height          =   495
         Index           =   11
         Left            =   2520
         Picture         =   "Form1.frx":3ABFA
         Stretch         =   -1  'True
         Tag             =   "ICEDL"
         ToolTipText     =   "E"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go DR"
         Height          =   495
         Index           =   10
         Left            =   720
         Picture         =   "Form1.frx":41758
         Stretch         =   -1  'True
         Tag             =   "ICEDR"
         ToolTipText     =   "E"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go U L"
         Height          =   495
         Index           =   9
         Left            =   120
         Picture         =   "Form1.frx":482B6
         Stretch         =   -1  'True
         Tag             =   "ICEUL"
         ToolTipText     =   "E"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice Go UR"
         Height          =   495
         Index           =   8
         Left            =   1920
         Picture         =   "Form1.frx":4EE14
         Stretch         =   -1  'True
         Tag             =   "ICEUR"
         ToolTipText     =   "E"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Ice"
         Height          =   495
         Index           =   7
         Left            =   1320
         Picture         =   "Form1.frx":55972
         Stretch         =   -1  'True
         Tag             =   "ICE"
         ToolTipText     =   "E"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Water"
         Height          =   495
         Index           =   6
         Left            =   720
         Picture         =   "Form1.frx":56AEC
         Stretch         =   -1  'True
         Tag             =   "WATER"
         ToolTipText     =   "E"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Fire"
         Height          =   495
         Index           =   5
         Left            =   120
         Picture         =   "Form1.frx":5757E
         Stretch         =   -1  'True
         Tag             =   "FIRE"
         ToolTipText     =   "E"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Block\Earth"
         Height          =   495
         Index           =   4
         Left            =   1320
         Picture         =   "Form1.frx":5EF76
         Stretch         =   -1  'True
         Tag             =   "H2"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Move Block"
         Height          =   495
         Index           =   3
         Left            =   1920
         Picture         =   "Form1.frx":5F7D8
         Stretch         =   -1  'True
         Tag             =   "ROCK"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Block"
         Height          =   495
         Index           =   2
         Left            =   2520
         Picture         =   "Form1.frx":60220
         Stretch         =   -1  'True
         Tag             =   "H"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Earth\Block"
         Height          =   495
         Index           =   1
         Left            =   720
         Picture         =   "Form1.frx":60E62
         Stretch         =   -1  'True
         Tag             =   "E2"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   495
      End
      Begin VB.Image OB 
         Appearance      =   0  'Flat
         DataSource      =   "Earth"
         Height          =   465
         Index           =   0
         Left            =   120
         Picture         =   "Form1.frx":6120E
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6465
      ScaleWidth      =   8625
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         Picture         =   "Form1.frx":61E50
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6495
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6495
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8655
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   0
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   8655
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   431
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   430
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   429
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   428
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   427
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   426
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   425
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   424
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   423
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   422
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   421
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   420
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   419
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   418
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   417
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   416
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   415
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   414
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   413
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   412
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   411
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   410
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   409
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   408
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   407
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   406
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   405
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   404
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   403
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   402
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   401
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   400
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   399
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   398
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   397
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   396
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   395
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   394
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   393
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   392
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   391
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   390
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   389
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   388
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   387
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   386
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   385
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   384
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   383
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   382
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   381
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   380
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   379
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   378
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   377
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   376
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   375
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   374
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   373
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   372
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   371
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   370
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   369
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   368
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   367
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   366
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   365
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   364
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   363
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   362
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   361
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   360
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   359
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   358
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   357
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   356
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   355
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   354
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   353
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   352
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   351
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   350
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   349
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   348
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   347
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   346
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   345
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   344
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   343
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   342
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   341
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   340
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   339
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   338
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   337
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   336
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   335
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   334
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   333
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   332
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   331
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   330
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   329
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   328
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   327
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   326
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   325
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   324
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   323
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   322
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   321
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   320
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   319
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   318
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   317
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   316
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   315
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   314
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   313
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   312
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   311
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   310
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   309
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   308
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   307
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   306
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   305
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   304
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   303
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   302
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   301
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   300
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   299
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   298
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   297
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   296
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   295
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   294
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   293
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   292
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   291
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   290
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   289
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   288
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   287
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   286
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   285
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   284
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   283
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   282
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   281
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   280
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   279
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   278
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   277
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   276
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   275
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   274
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   273
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   272
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   271
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   270
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   269
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   268
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   267
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   266
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   265
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   264
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   263
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   262
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   261
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   260
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   259
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   258
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   257
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   256
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   255
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   254
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   253
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   252
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   251
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   250
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   249
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   248
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   247
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   246
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   245
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   244
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   243
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   242
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   241
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   240
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   239
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   238
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   237
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   236
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   235
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   234
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   233
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   232
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   231
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   230
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   229
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   228
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   227
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   226
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   225
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   224
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   223
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   222
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   221
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   220
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   219
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   218
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   217
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   216
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   215
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   214
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   213
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   212
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   211
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   210
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   209
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   208
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   207
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   206
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   205
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   204
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   203
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   202
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   201
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   200
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   199
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   198
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   197
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   196
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   195
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   194
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   193
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   192
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   191
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   190
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   189
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   188
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   187
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   186
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   185
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   184
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   183
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   182
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   181
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   180
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   179
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   178
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   177
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   176
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   175
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   174
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   173
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   172
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   171
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   170
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   169
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   168
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   167
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   166
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   165
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   164
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   163
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   162
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   161
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   160
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   159
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   158
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   157
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   156
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   155
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   154
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   153
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   152
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   151
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   150
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   149
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   148
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   147
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   146
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   145
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   144
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   143
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   142
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   141
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   140
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   139
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   138
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   137
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   136
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   135
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   134
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   133
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   132
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   131
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   130
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   129
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   128
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   127
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   126
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   125
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   124
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   123
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   122
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   121
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   120
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   119
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   118
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   117
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   116
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   115
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   114
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   113
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   112
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   111
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   110
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   109
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   108
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   107
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   106
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   105
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   104
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   103
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   102
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   101
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   100
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   99
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   98
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   97
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   96
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   95
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   94
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   93
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   92
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   91
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   90
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   89
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   88
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   87
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   86
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   85
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   84
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   83
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   82
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   81
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   80
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   79
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   78
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   77
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   76
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   75
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   74
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   73
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   72
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   71
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   70
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   69
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   68
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   67
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   66
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   65
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   64
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   63
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   62
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   61
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   60
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   59
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   58
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   57
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   56
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   55
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   54
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   53
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   52
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   51
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   50
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   49
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   48
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   47
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   46
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   45
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   44
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   43
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   42
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   41
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   40
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   39
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   38
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   37
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   36
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   35
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   34
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   33
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   32
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   31
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   30
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   29
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   28
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   27
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   26
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   25
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   24
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   23
         Left            =   8280
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   22
         Left            =   7920
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   21
         Left            =   7560
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   20
         Left            =   7200
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   19
         Left            =   6840
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   18
         Left            =   6480
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   17
         Left            =   6120
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   16
         Left            =   5760
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   15
         Left            =   5400
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   14
         Left            =   5040
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   13
         Left            =   4680
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   12
         Left            =   4320
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   11
         Left            =   3960
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   10
         Left            =   3600
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   9
         Left            =   3240
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   8
         Left            =   2880
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   7
         Left            =   2520
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   6
         Left            =   2160
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   5
         Left            =   1800
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   4
         Left            =   1440
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   3
         Left            =   1080
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   2
         Left            =   720
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image E 
         Height          =   375
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "E"
         ToolTipText     =   "E"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Block : 000"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
For i = 0 To 999
LVN.List(i) = i
Next

On Error Resume Next
For i = 0 To 431
E(i).Tag = "E"
E(i).ToolTipText = "G"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(i).ToolTipText + ".bmp")
E(i).Appearance = 0
E(i).BorderStyle = 1
Next

On Error Resume Next
For i = 408 To 431
E(i).Tag = "H"
E(i).ToolTipText = "E"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\BLOCK" + E(i).ToolTipText + ".bmp")
Next

On Error Resume Next
For i = 0 To 23
E(i).Tag = "H"
E(i).ToolTipText = "E"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\BLOCK" + E(i).ToolTipText + ".bmp")
Next

On Error Resume Next
For i = 24 To 408
E(i).Tag = "H"
E(i).ToolTipText = "E"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\BLOCK" + E(i).ToolTipText + ".bmp")
i = i + 23
Next

On Error Resume Next
For i = 23 To 431
E(i).Tag = "H"
E(i).ToolTipText = "E"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\BLOCK" + E(i).ToolTipText + ".bmp")
i = i + 23
Next
E(75).Tag = "E"
E(75).ToolTipText = "E"
E(75).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(i).ToolTipText + ".bmp")
End Sub

Private Sub Command10_Click()
Form4.Show
End Sub

Private Sub Command11_Click()
List1.Clear
End Sub

Private Sub Command2_Click()
Dim X As Integer
X = 0
For i = 0 To 431
If E(i).Tag = "DH" Then
X = X + 1
End If
Next
CRI.Text = X
End Sub

Private Sub Command3_Click()
On Error GoTo 5
Dim i2 As Integer
NUM = FreeFile
Open App.path + "\Data Game\Levels\" + LVN.Text + ".Zan" For Input As NUM
Dim X As String
Input #NUM, X
PW.Text = X
Input #NUM, X
CRI.Text = X
For i = 0 To 431
Input #NUM, X
E(i).Tag = X
Input #NUM, X
E(i).ToolTipText = X
Next
Input #NUM, X
Combo1.Text = X
Input #NUM, i2
List1.Clear
For i = 1 To i2
Input #NUM, X
List1.AddItem X
Next
GoTo 6
5:
Close NUM
MsgBox "Sorry !! Load Level Not Complete !!!"
6:
Call Command5_Click
End Sub

Private Sub Command4_Click()
If Not PW.Text = "Enter PassWord" Then

Call Command2_Click

On Error GoTo 5
NUM = FreeFile
Open App.path + "\Data Game\Levels\" + LVN.Text + ".Zan" For Output As NUM
Write #NUM, PW.Text
Write #NUM, CRI.Text
For i = 0 To 431
Write #NUM, E(i).Tag
Write #NUM, E(i).ToolTipText
Next
Write #NUM, Combo1.Text
Write #NUM, List1.ListCount
For i = 0 To List1.ListCount - 1
Write #NUM, List1.List(i)
Next
Close NUM
MsgBox "Great !! Save Level Complete !!!"
GoTo 6
5:
MsgBox "Sorry !! Save Level Not Complete !!!"
6:
Else
MsgBox "Error !!!. Enter PassWord First Then Save The Level !"
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
For i = 0 To 431

If Not Right(GetFileName(E(i).Tag), 1) = "#" Then
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\" + GetFileName(E(i).Tag) + E(i).ToolTipText + ".bmp")
Else
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\" + Left(GetFileName(E(i).Tag), Len(GetFileName(E(i).Tag)) - 1) + ".bmp")
End If
Next
End Sub

Private Sub Command6_Click()
On Error Resume Next
Open App.path + "\Data Game\Save.ZAN" For Output As #1
Write #1, LVN.Text
Close #1
Shell App.path + "\PLAY.exe", vbNormalFocus
End Sub

Private Sub Command7_Click()
For i = 0 To 999
LVN.List(i) = i
Next

For i = 408 To 431
E(i).Tag = TA.Caption
E(i).ToolTipText = TOL.Caption
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\" + GetFileName(E(i).Tag) + E(i).ToolTipText + ".bmp")
Next

For i = 0 To 23
E(i).Tag = TA.Caption
E(i).ToolTipText = TOL.Caption
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\" + GetFileName(E(i).Tag) + E(i).ToolTipText + ".bmp")
Next

For i = 24 To 408
E(i).Tag = TA.Caption
E(i).ToolTipText = TOL.Caption
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\" + GetFileName(E(i).Tag) + E(i).ToolTipText + ".bmp")
i = i + 23
Next

For i = 23 To 431
E(i).Tag = TA.Caption
E(i).ToolTipText = TOL.Caption
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\" + GetFileName(E(i).Tag) + E(i).ToolTipText + ".bmp")
i = i + 23
Next
End Sub

Private Sub Command8_Click()
Load_Music (Int(Combo1.Text))
PlayMusic
SetMusic (100)
End Sub

Private Sub Command9_Click()
StopMusic
End Sub

Private Sub E_Click(Index As Integer)
E(Index).Tag = TA.Caption
E(Index).ToolTipText = TOL.Caption
E(Index).Picture = PS.Picture
End Sub

Private Sub E_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Block : " + Str(E(Index).Index)
End Sub

Private Sub Form_Load()

Initialize_Music (25)

Call Command1_Click

Combo1.Text = Combo1.List(0)

Open App.path + "\Data Game\View.ZAN" For Output As #1
Write #1, "ON"
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open App.path + "\Data Game\View.ZAN" For Output As #1
Write #1, "OFF"
Close #1
End
End Sub

Private Sub HLP_Change()
OB(30).ToolTipText = HLP.Text
End Sub

Private Sub Image1_Click()
Command7_Click
End Sub

Private Sub Image2_Click()
Command7_Click
End Sub

Private Sub Image3_Click()
Command7_Click
End Sub

Private Sub Image4_Click()
Command7_Click
End Sub

Private Sub Image5_Click()
E_Click (75)
End Sub

Private Sub OB_Click(Index As Integer)
NAM.Caption = HL.Caption
PS.Picture = OB(Index).Picture
TA.Caption = OB(Index).Tag
TOL.Caption = OB(Index).ToolTipText

If Not GetFileName(OB(Index).Tag) = "" Then
If Not Right(GetFileName(OB(Index).Tag), 1) = "#" Then
Form3.Caption = GetFileName(OB(Index).Tag)
Form3.Show 1
Else
Form3.Caption = Left(GetFileName(OB(Index).Tag), Len(GetFileName(OB(Index).Tag)) - 1)
End If
End If
End Sub

Private Sub OB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case OB(Index).Index
Case Is = 0: HL.Caption = "Earth"
Case Is = 1: HL.Caption = "Earth\Block"
Case Is = 2: HL.Caption = "Block"
Case Is = 3: HL.Caption = "Move Block"
Case Is = 4: HL.Caption = "Block\Earth"
Case Is = 5: HL.Caption = "Fire"
Case Is = 6: HL.Caption = "Water"
Case Is = 7: HL.Caption = "Ice"
Case Is = 8: HL.Caption = "Ice Go UR"
Case Is = 9: HL.Caption = "Ice Go UL"
Case Is = 10: HL.Caption = "Ice Go DR"
Case Is = 11: HL.Caption = "Ice Go DL"
Case Is = 12: HL.Caption = "Ice Go RU"
Case Is = 13: HL.Caption = "Ice Go RD"
Case Is = 14: HL.Caption = "Ice Go LU"
Case Is = 15: HL.Caption = "Ice Go LD"
Case Is = 16: HL.Caption = "Button"
Case Is = 17: HL.Caption = "Have Red Key"
Case Is = 18: HL.Caption = "Have Blue Key"
Case Is = 19: HL.Caption = "Have Yellow Key"
Case Is = 20: HL.Caption = "Red Key"
Case Is = 21: HL.Caption = "Blue Key"
Case Is = 22: HL.Caption = "Yellow Key"
Case Is = 23: HL.Caption = "Fire Boots"
Case Is = 24: HL.Caption = "Water Boots"
Case Is = 25: HL.Caption = "Ice Boots"
Case Is = 26: HL.Caption = "Com U"
Case Is = 27: HL.Caption = "Com R"
Case Is = 28: HL.Caption = "Com D"
Case Is = 29: HL.Caption = "Com L"
Case Is = 30: HL.Caption = "Tips Block"
Case Is = 31: HL.Caption = "End Game"
Case Is = 32: HL.Caption = "Get 1 Cristal"
End Select
End Sub

Private Sub PW_Click()
If PW.Text = "Enter PassWord" Then
PW.Text = ""
End If
End Sub


Function GetFileName(Code As String) As String
If Code = "E" Then
GetFileName = "E"
End If
If Code = "DH" Then
GetFileName = "DH"
End If
If Code = "END" Then
GetFileName = "END"
End If
If Code = "GETFIRE" Then
GetFileName = "GETFIRE"
End If
If Code = "LU" Then
GetFileName = "LU"
End If
If Code = "H2" Then
GetFileName = "H2"
End If
If Code = "E2" Then
GetFileName = "E2"
End If
If Code = "GETWATER" Then
GetFileName = "GETWATER"
End If
If Code = "GETICE" Then
GetFileName = "GETICE"
End If
If Code = "FIRE" Then
GetFileName = "FIRE"
End If
If Code = "TIPS" Then
GetFileName = "TIPS#"
End If
If Code = "ICE" Then
GetFileName = "ICE"
End If
If Code = "H" Then
GetFileName = "BLOCK"
End If
If Code = "ROCK" Then
GetFileName = "ROCK"
End If
If Code = "WATER" Then
GetFileName = "WATER"
End If
If Code = "ICEU" Then
GetFileName = "ICEU"
End If
If Code = "ICED" Then
GetFileName = "ICED"
End If
If Code = "ICEL" Then
GetFileName = "ICEL"
End If
If Code = "ICER" Then
GetFileName = "ICER"
End If
If Code = "ICEUR" Then
GetFileName = "ICEUR"
End If
If Code = "ICEUL" Then
GetFileName = "ICEUL"
End If
If Code = "ICEDR" Then
GetFileName = "ICEDR"
End If
If Code = "ICEDL" Then
GetFileName = "ICEDL"
End If
If Code = "ICERU" Then
GetFileName = "ICERU"
End If
If Code = "ICERD" Then
GetFileName = "ICERD"
End If
If Code = "ICELU" Then
GetFileName = "ICELU"
End If
If Code = "ICELD" Then
GetFileName = "ICELD"
End If
If Code = "GETRED" Then
GetFileName = "GR"
End If
If Code = "GETBLUE" Then
GetFileName = "GB"
End If
If Code = "GETYELLOW" Then
GetFileName = "GY"
End If
If Code = "DOORHR" Then
GetFileName = "HKR#"
End If
If Code = "DOORHB" Then
GetFileName = "HKB#"
End If
If Code = "DOORHY" Then
GetFileName = "HKY#"
End If
If Code = "COM" Then
GetFileName = "COM"
End If
End Function

