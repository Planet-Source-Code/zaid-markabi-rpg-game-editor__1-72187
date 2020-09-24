VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ÇäåÇÁ ÇááÚÈÉ \ ÇÞÊÑÇÝ.ÞÇÑÝ"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6405
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   960
      Top             =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YOU WIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   67.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   6255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
On Error GoTo 5
NUM = FreeFile
Open App.path + "\Data Game\Save.ZAN" For Output As NUM
Write #NUM, "0"
Close NUM
Form1.Show
5:
Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Form1
Timer1.Interval = 0
End Sub
