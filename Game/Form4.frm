VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ÇÞÊÑÇÝ.ÞÇÑÝ - Get Wrong"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7185
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by  Zaid Markabi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4800
      TabIndex        =   3
      Top             =   3360
      Width           =   2310
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "Form4.frx":0CCA
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   960
      Index           =   1
      Left            =   0
      Picture         =   "Form4.frx":238C
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1920
      Index           =   0
      Left            =   0
      Picture         =   "Form4.frx":7DCE
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   360
      Left            =   4320
      TabIndex        =   1
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   360
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "Form4.frx":1E610
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImgChng As Boolean
Dim ImgChngFrm As Integer

Private Sub Form_Load()
Me.Picture = Image1.Picture

Dim x As String
Open App.path + "\Data Game\View.ZAN" For Input As #1
Input #1, x
Close #1

Open App.path + "\Data Game\View.ZAN" For Output As #2
Write #2, "OFF"
Close #2

If x = "ON" Then
Me.Show
Label2_Click
End If
End Sub

Private Sub Label2_Click()
On Error Resume Next
Unload Me
Form1.Show
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_Click()
On Error GoTo 5
NUM = FreeFile
Open App.path + "\Data Game\Save.ZAN" For Output As NUM
Write #NUM, "0"
Close NUM
Unload Me
Form1.Show
5:
End Sub

Private Sub Timer1_Timer()

If ImgChng = True Then
ImgChngFrm = ImgChngFrm + 1
Else
ImgChngFrm = ImgChngFrm - 1
End If

If ImgChngFrm < 0 Or ImgChngFrm > 1 Then ImgChng = Not ImgChng

If ImgChngFrm = -1 Then
Image1.Picture = Me.Picture
Timer1.Interval = 500
Else
Image1.Picture = Image2(ImgChngFrm).Picture
Timer1.Interval = 100
End If
End Sub
