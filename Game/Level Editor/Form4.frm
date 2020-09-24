VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   LinkTopic       =   "Form4"
   ScaleHeight     =   3855
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add to the Story"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   120
      List            =   "Form4.frx":0025
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Skip"
         Height          =   255
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00808080&
         Caption         =   "Next"
         Height          =   255
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label TextTalking 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label NameTalk 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   120
         Width           =   3015
      End
      Begin VB.Image WhoTalk 
         Height          =   2655
         Left            =   0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2415
      End
      Begin VB.Image BackStory 
         Height          =   2775
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
NameTalk.Caption = Combo1.Text + " :"
WhoTalk.Picture = LoadPicture(App.path + "\Data Game\Story\" + Combo1.Text + ".emf")
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Command2_Click()
Form1.List1.AddItem Combo1.Text
Form1.List1.AddItem Text1.Text
End Sub

Private Sub Form_Load()
BackStory.Picture = LoadPicture(App.path + "\Data Game\Story\Back.jpg")
Combo1.Text = Combo1.List(0)
End Sub

Private Sub Text1_Change()
TextTalking.Caption = Text1.Text
End Sub
