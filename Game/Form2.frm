VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter Password"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.TextBox PW 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox LV 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "PassWord :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Level Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo 5
NUM = FreeFile
Open App.Path + "\Data Game\Levels\" + LV.Text + ".Zan" For Input As NUM
Dim X As String
Input #NUM, X
Close NUM
If X = PW.Text Then
Unload Form1
NUM = FreeFile
Open App.Path + "\Data Game\Save.ZAN" For Output As NUM
Write #NUM, LV.Text
Close NUM
Form1.Show
Unload Me
End If
5:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
For i = 0 To 999
LV.List(i) = i
Next
LV.Text = "0"
End Sub
