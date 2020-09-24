VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Text            =   "c:\NewImage.Bmp"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   720
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2160
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "#"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Image :"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Item :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Add New Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function IsItemLocated(File As String) As Boolean
On Error GoTo 1
Picture1.Picture = LoadPicture(App.path + "\Data Game\Item\" + File + ".bmp")
IsItemLocated = True
Exit Function
1:
IsItemLocated = False
End Function

Function NextName(NameI As String) As String
Dim ListName As String
ListName = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
Dim n As Integer
n = InStr(1, ListName, NameI)
NextName = Mid(ListName, n + 1, 1)
End Function

Function NewItemName() As String
Dim TryName As String
TryName = "A"

Do While IsItemLocated(Me.Caption + TryName) = True And Not TryName = "9"
TryName = NextName(TryName)
Loop

NewItemName = TryName
End Function

Private Sub Command1_Click()
If Not Picture2.Picture = Me.Picture Then
SavePicture Picture2.Picture, App.path + "\Data Game\Item\" + Text1.Text + ".bmp"
Unload Me
Unload Form3
Else
MsgBox "Enter URL of new Image !", , "File Path"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = Form3.Caption

If Not NewItemName = "9" Then
Text1.Text = Me.Caption + NewItemName
Else
MsgBox "Can't add more items in this part !"
Unload Me
End If

End Sub

Private Sub Text2_Change()
On Error Resume Next
Picture2.Picture = LoadPicture(Text2.Text)
End Sub
