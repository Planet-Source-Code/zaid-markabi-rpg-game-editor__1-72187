VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   360
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Select your Item"
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
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   65
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   64
      Left            =   0
      Picture         =   "Form3.frx":03AC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   63
      Left            =   0
      Picture         =   "Form3.frx":0758
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   62
      Left            =   0
      Picture         =   "Form3.frx":0B04
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   61
      Left            =   0
      Picture         =   "Form3.frx":0EB0
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   60
      Left            =   0
      Picture         =   "Form3.frx":125C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   59
      Left            =   0
      Picture         =   "Form3.frx":1608
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   58
      Left            =   0
      Picture         =   "Form3.frx":19B4
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   57
      Left            =   0
      Picture         =   "Form3.frx":1D60
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   56
      Left            =   1200
      Picture         =   "Form3.frx":210C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   55
      Left            =   600
      Picture         =   "Form3.frx":24B8
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   54
      Left            =   600
      Picture         =   "Form3.frx":2864
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   53
      Left            =   1200
      Picture         =   "Form3.frx":2C10
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   52
      Left            =   0
      Picture         =   "Form3.frx":2FBC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   51
      Left            =   0
      Picture         =   "Form3.frx":3368
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   50
      Left            =   0
      Picture         =   "Form3.frx":3714
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   49
      Left            =   0
      Picture         =   "Form3.frx":3AC0
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   48
      Left            =   0
      Picture         =   "Form3.frx":3E6C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   47
      Left            =   0
      Picture         =   "Form3.frx":4218
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   46
      Left            =   0
      Picture         =   "Form3.frx":45C4
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   45
      Left            =   0
      Picture         =   "Form3.frx":4970
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   44
      Left            =   0
      Picture         =   "Form3.frx":4D1C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   43
      Left            =   0
      Picture         =   "Form3.frx":50C8
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   42
      Left            =   0
      Picture         =   "Form3.frx":5474
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   41
      Left            =   0
      Picture         =   "Form3.frx":5820
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   40
      Left            =   0
      Picture         =   "Form3.frx":5BCC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   39
      Left            =   0
      Picture         =   "Form3.frx":5F78
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   38
      Left            =   0
      Picture         =   "Form3.frx":6324
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   37
      Left            =   0
      Picture         =   "Form3.frx":66D0
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   36
      Left            =   0
      Picture         =   "Form3.frx":6A7C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   35
      Left            =   0
      Picture         =   "Form3.frx":6E28
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   34
      Left            =   1200
      Picture         =   "Form3.frx":71D4
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   33
      Left            =   600
      Picture         =   "Form3.frx":7580
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   32
      Left            =   600
      Picture         =   "Form3.frx":792C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   31
      Left            =   1200
      Picture         =   "Form3.frx":7CD8
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   30
      Left            =   0
      Picture         =   "Form3.frx":8084
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   29
      Left            =   0
      Picture         =   "Form3.frx":8430
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   28
      Left            =   0
      Picture         =   "Form3.frx":87DC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   27
      Left            =   0
      Picture         =   "Form3.frx":8B88
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   26
      Left            =   0
      Picture         =   "Form3.frx":8F34
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   25
      Left            =   0
      Picture         =   "Form3.frx":92E0
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   24
      Left            =   0
      Picture         =   "Form3.frx":968C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   23
      Left            =   0
      Picture         =   "Form3.frx":9A38
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   22
      Left            =   0
      Picture         =   "Form3.frx":9DE4
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   21
      Left            =   0
      Picture         =   "Form3.frx":A190
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   20
      Left            =   0
      Picture         =   "Form3.frx":A53C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   19
      Left            =   0
      Picture         =   "Form3.frx":A8E8
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   18
      Left            =   0
      Picture         =   "Form3.frx":AC94
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   17
      Left            =   0
      Picture         =   "Form3.frx":B040
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   16
      Left            =   0
      Picture         =   "Form3.frx":B3EC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   15
      Left            =   0
      Picture         =   "Form3.frx":B798
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   14
      Left            =   0
      Picture         =   "Form3.frx":BB44
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   13
      Left            =   0
      Picture         =   "Form3.frx":BEF0
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   12
      Left            =   1200
      Picture         =   "Form3.frx":C29C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   11
      Left            =   600
      Picture         =   "Form3.frx":C648
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   10
      Left            =   600
      Picture         =   "Form3.frx":C9F4
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   9
      Left            =   1200
      Picture         =   "Form3.frx":CDA0
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   8
      Left            =   0
      Picture         =   "Form3.frx":D14C
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   7
      Left            =   0
      Picture         =   "Form3.frx":D4F8
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   6
      Left            =   0
      Picture         =   "Form3.frx":D8A4
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   5
      Left            =   0
      Picture         =   "Form3.frx":DC50
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   4
      Left            =   0
      Picture         =   "Form3.frx":DFFC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   3
      Left            =   0
      Picture         =   "Form3.frx":E3A8
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   2
      Left            =   0
      Picture         =   "Form3.frx":E754
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   1
      Left            =   0
      Picture         =   "Form3.frx":EB00
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image OB 
      Appearance      =   0  'Flat
      DataSource      =   "Earth"
      Height          =   585
      Index           =   0
      Left            =   0
      Picture         =   "Form3.frx":EEAC
      Stretch         =   -1  'True
      Tag             =   "E"
      ToolTipText     =   "E"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub RefreshIcons()
Dim RowNum As Integer
RowNum = 4

Dim i, ObjTop, ObjLeft As Integer

Dim ObjNum As Integer
For i = 0 To OB.Count - 1
If OB(i).Visible = True Then ObjNum = ObjNum + 1
Next

ObjTop = OB(0).Top

For i = 1 To ObjNum

ObjLeft = ObjLeft + 1
If ObjLeft = RowNum Then
ObjLeft = 0
ObjTop = ObjTop + OB(0).Height
End If

OB(i).Left = ObjLeft * OB(i).Width
OB(i).Top = ObjTop

Next

OB(ObjNum).Visible = True
OB(ObjNum).Picture = LoadPicture(App.path + "\Data Game\Item\NEW.bmp")
OB(ObjNum).ToolTipText = "Add New Block"

Me.Height = ObjTop + OB(0).Height
Me.Width = OB(0).Width * RowNum
Me.Top = Me.Top - OB(0).Width * RowNum
End Sub

Private Sub Form_Load()
Me.Width = 0
Me.Height = 0
End Sub

Private Sub Label1_Click(Index As Integer)
Unload Me
End Sub

Private Sub OB_Click(Index As Integer)
On Error Resume Next
If Not OB(Index).ToolTipText = "Add New Block" Then
Form1.TOL.Caption = OB(Index).ToolTipText
Form1.PS.Picture = OB(Index).Picture
Unload Me
Else
Form2.Show 1
End If
End Sub

Private Sub Timer1_Timer()
File1.path = App.path + "\Data Game\Item\"

Dim ObjNum As Integer

For i = 0 To File1.ListCount - 1
If Left(File1.List(i), Len(Me.Caption)) = Me.Caption And Len(File1.List(i)) = Len(Me.Caption) + 5 Then
OB(ObjNum).Picture = LoadPicture(App.path + "\Data Game\Item\" + File1.List(i))
OB(ObjNum).ToolTipText = Mid(File1.List(i), Len(Me.Caption) + 1, 1)
OB(ObjNum).Visible = True
ObjNum = ObjNum + 1
End If
Next
RefreshIcons
Timer1.Enabled = False
'If ObjNum = 2 Then OB_Click (0)
End Sub
