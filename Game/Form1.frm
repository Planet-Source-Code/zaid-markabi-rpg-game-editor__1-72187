VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Get Wrong"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleMode       =   0  'User
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer DACOM 
      Interval        =   250
      Left            =   2040
      Top             =   1680
   End
   Begin VB.Timer DAROI 
      Interval        =   1
      Left            =   2040
      Top             =   1200
   End
   Begin VB.Timer DAON 
      Left            =   2160
      Top             =   2880
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   5415
      TabIndex        =   43
      Top             =   360
      Width           =   5415
      Begin VB.CommandButton Command4 
         BackColor       =   &H00808080&
         Caption         =   "Skip"
         Height          =   255
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00808080&
         Caption         =   "Next"
         Height          =   255
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2400
         Width           =   495
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
         TabIndex        =   45
         Top             =   120
         Width           =   3015
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
         TabIndex        =   44
         Top             =   480
         Width           =   2775
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
   Begin VB.PictureBox Picture6 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   1275
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Label Label2 
         Caption         =   "FIRE"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   375
      End
      Begin VB.Label FI 
         Caption         =   "OFF"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "WAT"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   375
      End
      Begin VB.Label WA 
         Caption         =   "OFF"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   240
         Width           =   375
      End
      Begin VB.Label IC 
         Caption         =   "OFF"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "ICE"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   375
      End
      Begin VB.Label HLP 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label KB 
         Caption         =   "OFF"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "BLU"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label KR 
         Caption         =   "OFF"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "RED"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   375
      End
      Begin VB.Label KY 
         Caption         =   "OFF"
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "YEL"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label DHB 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label ALLDHB 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label LV 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   855
      End
      Begin WMPLibCtl.WindowsMediaPlayer S 
         Height          =   975
         Left            =   0
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   2143
         _cy             =   1720
      End
   End
   Begin VB.TextBox EN 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "75"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox CON 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame CH 
      BackColor       =   &H00000000&
      Caption         =   "Cheat Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "Cancle"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Enter"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox CHT 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Text            =   "Enter Password"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5295
      Left            =   3600
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   17
         Top             =   4800
         Width           =   1695
         Begin VB.Label PW 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   14
         Top             =   4080
         Width           =   1695
         Begin VB.Label LVL 
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   1635
         TabIndex        =   12
         Top             =   2880
         Width           =   1695
         Begin VB.Label HL 
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   735
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
         Begin VB.Label DHH 
            BackColor       =   &H00000000&
            Caption         =   "0\0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Level :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tips :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cristal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Keys :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Boots :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Image DEL 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image C1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image B1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image A1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image C 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   480
         Width           =   495
      End
      Begin VB.Image B 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   720
         Stretch         =   -1  'True
         Top             =   480
         Width           =   495
      End
      Begin VB.Image A 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.TextBox DI 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "U"
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3345
      ScaleWidth      =   3345
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.PictureBox L 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   9015
         Left            =   0
         ScaleHeight     =   9015
         ScaleMode       =   0  'User
         ScaleWidth      =   11895
         TabIndex        =   0
         Top             =   0
         Width           =   11895
         Begin VB.Timer Timer1 
            Interval        =   5000
            Left            =   2040
            Top             =   120
         End
         Begin VB.Image Y 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1440
            Picture         =   "Form1.frx":0000
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   431
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   430
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   429
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   428
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   427
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   426
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   425
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   424
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   423
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   422
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   421
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   420
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   419
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   418
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   417
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   416
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   415
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   414
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   413
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   412
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   411
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   410
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   409
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   408
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   8160
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   407
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   406
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   405
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   404
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   403
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   402
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   401
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   400
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   399
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   398
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   397
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   396
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   395
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   394
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   393
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   392
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   391
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   390
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   389
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   388
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   387
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   386
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   385
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   384
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7680
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   383
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   382
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   381
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   380
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   379
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   378
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   377
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   376
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   375
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   374
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   373
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   372
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   371
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   370
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   369
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   368
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   367
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   366
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   365
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   364
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   363
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   362
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   361
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   360
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   7200
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   359
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   358
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   357
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   356
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   355
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   354
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   353
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   352
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   351
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   350
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   349
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   348
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   347
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   346
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   345
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   344
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   343
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   342
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   341
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   340
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   339
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   338
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   337
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   336
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   335
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   334
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   333
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   332
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   331
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   330
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   329
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   328
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   327
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   326
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   325
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   324
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   323
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   322
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   321
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   320
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   319
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   318
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   317
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   316
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   315
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   314
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   313
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   312
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   6240
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   311
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   310
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   309
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   308
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   307
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   306
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   305
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   304
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   303
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   302
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   301
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   300
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   299
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   298
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   297
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   296
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   295
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   294
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   293
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   292
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   291
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   290
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   289
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   288
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5760
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   287
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   286
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   285
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   284
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   283
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   282
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   281
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   280
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   279
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   278
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   277
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   276
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   275
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   274
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   273
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   272
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   271
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   270
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   269
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   268
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   267
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   266
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   265
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   264
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   5280
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   263
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   262
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   261
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   260
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   259
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   258
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   257
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   256
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   255
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   254
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   253
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   252
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   251
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   250
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   249
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   248
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   247
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   246
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   245
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   244
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   243
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   242
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   241
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   240
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4800
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   239
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   238
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   237
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   236
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   235
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   234
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   233
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   232
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   231
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   230
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   229
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   228
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   227
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   226
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   225
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   224
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   223
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   222
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   221
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   220
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   219
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   218
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   217
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   216
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   215
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   214
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   213
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   212
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   211
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   210
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   209
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   208
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   207
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   206
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   205
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   204
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   203
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   202
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   201
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   200
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   199
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   198
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   197
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   196
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   195
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   194
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   193
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   192
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3840
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   191
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   190
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   189
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   188
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   187
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   186
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   185
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   184
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   183
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   182
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   181
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   180
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   179
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   178
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   177
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   176
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   175
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "ICERU"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   174
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "ICE"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   173
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "ICE"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   172
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "ICE"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   171
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "ICE"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   170
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "ICE"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   169
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "ICE"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   168
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "ICELU"
            Top             =   3360
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   167
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   166
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   165
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   164
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   163
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   162
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   161
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   160
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   159
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   158
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   157
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   156
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   155
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   154
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   153
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   152
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   151
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "ICED"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   150
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E2"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   149
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   148
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   147
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   146
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   145
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   144
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   143
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   142
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   141
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   140
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   139
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   138
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   137
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   136
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   135
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   134
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   133
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   132
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   131
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   130
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   129
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   128
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   127
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   126
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   125
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   124
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   123
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   122
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   121
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   120
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   119
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   118
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   117
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   116
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   115
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   114
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   113
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   112
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   111
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   110
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   109
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   108
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   107
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   106
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   105
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   104
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   103
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   102
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "ROCK"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   101
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   100
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   99
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "DH"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   98
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   97
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "COM"
            ToolTipText     =   "R"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   96
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   95
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   94
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   93
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   92
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   91
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   90
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   89
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   88
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   87
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   86
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   85
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   84
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   83
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   82
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   81
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   80
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   79
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   78
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   77
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   76
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "LU"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   75
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   74
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   73
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "GETICE"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   72
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   71
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   70
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   69
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   68
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   67
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   66
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   65
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   64
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   63
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   62
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   61
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   60
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   59
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   58
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   57
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   56
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   55
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   54
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   53
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   52
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   51
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   50
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   49
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   48
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   960
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   47
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   46
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   45
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   44
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   43
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   42
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   41
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   40
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   39
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   38
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   37
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   36
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   35
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   34
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   33
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   32
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   31
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   30
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   29
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   28
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   27
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   26
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   25
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "WATER"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   24
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   480
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   23
            Left            =   11040
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   22
            Left            =   10560
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   21
            Left            =   10080
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   20
            Left            =   9600
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   19
            Left            =   9120
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   18
            Left            =   8640
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   17
            Left            =   8160
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   16
            Left            =   7680
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   15
            Left            =   7200
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   14
            Left            =   6720
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   13
            Left            =   6240
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   12
            Left            =   5760
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   11
            Left            =   5280
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   10
            Left            =   4800
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   9
            Left            =   4320
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   8
            Left            =   3840
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   7
            Left            =   3360
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   6
            Left            =   2880
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   5
            Left            =   2400
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   4
            Left            =   1920
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   3
            Left            =   1440
            Stretch         =   -1  'True
            Tag             =   "H2"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   2
            Left            =   960
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   1
            Left            =   480
            Stretch         =   -1  'True
            Tag             =   "E2"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image E 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Tag             =   "E"
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Game"
      End
      Begin VB.Menu GSFS 
         Caption         =   "Level Password"
      End
      Begin VB.Menu mnuCheats 
         Caption         =   "Cheats"
      End
      Begin VB.Menu FFAE 
         Caption         =   "Exit Game"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "Setting"
      Begin VB.Menu mnuSize 
         Caption         =   "Map Size"
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu mnuSmall 
            Caption         =   "Small"
         End
         Begin VB.Menu mnuLarge 
            Caption         =   "Large"
         End
      End
      Begin VB.Menu mnuSounds 
         Caption         =   "Sounds \ Music"
         Begin VB.Menu mnuMusic 
            Caption         =   "Music"
            Begin VB.Menu mnuMuteMus 
               Caption         =   "Mute"
            End
            Begin VB.Menu mnuVol1 
               Caption         =   "50%"
            End
            Begin VB.Menu mnuVol2 
               Caption         =   "100%"
            End
            Begin VB.Menu mnuVol3 
               Caption         =   "200%"
            End
         End
         Begin VB.Menu mnuSound 
            Caption         =   "Sounds"
            Begin VB.Menu mnuVolOn 
               Caption         =   "On"
            End
            Begin VB.Menu mnuVolOff 
               Caption         =   "Off"
            End
         End
      End
   End
   Begin VB.Menu mnuAboutGm 
      Caption         =   "About"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Get-Wrong"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Get Wrong  - VB 6.0 game
' Written by  Zaid Markabi
'
' for more VB6 Games , visit my website
' Http://www.yazanmarkabi.webs.com
'
' Email : ZaidMarkabi@yahoo.com

Dim PlayerPos As Integer
Dim PressWait As Integer
Dim StoryNum As Integer
Dim NextStory As Boolean
Dim SkipStory As Boolean
Dim EndGame As Boolean

Private Sub ALLDHB_Change()
DHH.Caption = DHB.Caption + " \ " + ALLDHB.Caption
End Sub

Private Sub CHT_LostFocus()
CH.Visible = False
End Sub

Private Sub Command1_Click()
If CHT.Text = "Next" Then
LV.Caption = LV.Caption + 1
Exit Sub
End If

If CHT.Text = "Back" Then
LV.Caption = LV.Caption - 1
Exit Sub
End If

If CHT.Text = "Item" Then
KR.Caption = "ON"
KB.Caption = "ON"
KY.Caption = "ON"
FI.Caption = "ON"
WA.Caption = "ON"
IC.Caption = "ON"
DHB.Caption = ALLDHB.Caption
HLP.Caption = "Have All Items"
Exit Sub
End If

MsgBox "Wrong Password !!!."
CH.Visible = False
End Sub

Private Sub Command2_Click()
CH.Visible = False
CON.SetFocus
End Sub

Private Sub Command3_Click()
NextStory = True
End Sub

Private Sub Command4_Click()
SkipStory = True
End Sub

Private Sub CON_Change()
DAON.Interval = 1

If CON.Text = "E" Then
CH.Visible = True
End If

s.URL = App.path + "\Data Game\Sounds\step.WAV"

If E(EN.Text).Tag = "DOORHB" And E(EN.Text).ToolTipText = "LOCKED" Then
If KB.Caption = "ON" Then
E(EN.Text).ToolTipText = "UNLOCKED"
Else
If DI.Text = "R" Then
CON.Text = "A"
GoTo 6
End If
If DI.Text = "L" Then
CON.Text = "D"
GoTo 6
End If
If DI.Text = "U" Then
CON.Text = "S"
GoTo 6
End If
If DI.Text = "D" Then
CON.Text = "W"
GoTo 6
End If
End If
End If

If E(EN.Text).Tag = "DOORHY" And E(EN.Text).ToolTipText = "LOCKED" Then
If KY.Caption = "ON" Then
E(EN.Text).ToolTipText = "UNLOCKED"
Else
If DI.Text = "R" Then
CON.Text = "A"
GoTo 6
End If
If DI.Text = "L" Then
CON.Text = "D"
GoTo 6
End If
If DI.Text = "U" Then
CON.Text = "S"
GoTo 6
End If
If DI.Text = "D" Then
CON.Text = "W"
GoTo 6
End If
End If
End If

If E(EN.Text).Tag = "DOORHR" And E(EN.Text).ToolTipText = "LOCKED" Then
If KR.Caption = "ON" Then
E(EN.Text).ToolTipText = "UNLOCKED"
Else
If DI.Text = "R" Then
CON.Text = "A"
GoTo 6
End If
If DI.Text = "L" Then
CON.Text = "D"
GoTo 6
End If
If DI.Text = "U" Then
CON.Text = "S"
GoTo 6
End If
If DI.Text = "D" Then
CON.Text = "W"
GoTo 6
End If
End If
End If

6:

If Not E(EN.Text).Tag = "WATER" Then
y.Picture = LoadPicture(App.path + "\Data Game\Item\Y" + DI.Text + ".bmp")
Else
y.Picture = LoadPicture(App.path + "\Data Game\Item\YW" + DI.Text + ".bmp")
End If

If CON.Text = "Q" Then
End
End If

If Not E(EN.Text + 1).Tag = "H" Then
If Not E(EN.Text + 1).Tag = "H2" Then
If CON.Text = "D" Or CON.Text = "d" Then
y.Left = y.Left + E(0).Width
L.Left = L.Left - E(0).Width
EN.Text = EN.Text + 1
DI.Text = "R"
CON.Text = ""
Exit Sub
End If
End If
End If

If Not E(EN.Text - 1).Tag = "H" Then
If Not E(EN.Text - 1).Tag = "H2" Then
If CON.Text = "A" Or CON.Text = "a" Then
y.Left = y.Left - E(0).Width
L.Left = L.Left + E(0).Width
EN.Text = EN.Text - 1
DI.Text = "L"
CON.Text = ""
Exit Sub
End If
End If
End If

If Not E(EN.Text - 24).Tag = "H" Then
If Not E(EN.Text - 24).Tag = "H2" Then
If CON.Text = "W" Or CON.Text = "w" Then
y.Top = y.Top - E(0).Width
L.Top = L.Top + E(0).Width
EN.Text = EN.Text - 24
DI.Text = "U"
CON.Text = ""
Exit Sub
End If
End If
End If

If Not E(EN.Text + 24).Tag = "H" Then
If Not E(EN.Text + 24).Tag = "H2" Then
If CON.Text = "S" Or CON.Text = "s" Then
y.Top = y.Top + E(0).Width
L.Top = L.Top - E(0).Width
EN.Text = EN.Text + 24
DI.Text = "D"
CON.Text = ""
Exit Sub
End If
End If
End If

CON.Text = ""
End Sub

Private Sub DACOM_Timer()
Dim ComMoved As String

For i = 0 To 431

If E(i).Tag = "COM" Then

If E(i).Index = EN.Text Then
MsgBox "You Died !!. Because You Walk On Zombie !!.", vbOKOnly + vbInformation, "You Died"
Unload Me
Me.Show
Exit Sub
End If

If InStr(1, ComMoved, "*" + Format(i) + "*") = 0 Then

If E(i).ToolTipText = "R" Then
If E(i + 1).Tag = "E" Then
E(i).Tag = "E"
E(i).ToolTipText = E(i + 1).DataField
E(i + 1).Tag = "COM"
E(i + 1).ToolTipText = "R"
ComMoved = ComMoved + "*" + Format(i + 1) + "*"
GoTo 1
Else
E(i).ToolTipText = "U"
GoTo 1
End If
End If

If E(i).ToolTipText = "U" Then
If E(i - 24).Tag = "E" Then
E(i).Tag = "E"
E(i).ToolTipText = E(i - 24).DataField
E(i - 24).Tag = "COM"
E(i - 24).ToolTipText = "U"
ComMoved = ComMoved + "*" + Format(i - 24) + "*"
GoTo 1
Else
E(i).ToolTipText = "L"
GoTo 1
End If
End If

If E(i).ToolTipText = "L" Then
If E(i - 1).Tag = "E" Then
E(i).Tag = "E"
E(i).ToolTipText = E(i - 1).DataField
E(i - 1).Tag = "COM"
E(i - 1).ToolTipText = "L"
ComMoved = ComMoved + "*" + Format(i - 1) + "*"
GoTo 1
Else
E(i).ToolTipText = "D"
GoTo 1
End If
End If

If E(i).ToolTipText = "D" Then
If E(i + 24).Tag = "E" Then
E(i).Tag = "E"
E(i).ToolTipText = E(i + 24).DataField
E(i + 24).Tag = "COM"
E(i + 24).ToolTipText = "D"
ComMoved = ComMoved + "*" + Format(i + 24) + "*"
GoTo 1
Else
E(i).ToolTipText = "R"
GoTo 1
End If
End If

End If

End If

1:
Next
End Sub

Private Sub DAON_Timer()
DAON.Interval = 0
If E(EN.Text).Tag = "E" Then
Exit Sub
End If

If E(EN.Text).Tag = "E2" Then
Exit Sub
End If

If E(EN.Text).Tag = "DH" Then
DHB.Caption = DHB.Caption + 1
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
s.URL = App.path + "\Data Game\Sounds\Gold.WAV"
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\" + E(Int(EN.Text)).Tag + E(Int(EN.Text)).ToolTipText + ".bmp")
Exit Sub
End If

If E(EN.Text).Tag = "END" Then
If DHB.Caption = ALLDHB.Caption Then
LV.Caption = LV.Caption + 1
Unload Me
Me.Show
Exit Sub
End If
End If

If E(EN.Text).Tag = "FIRE" Then
If FI.Caption = "OFF" Then
s.URL = App.path + "\Data Game\Sounds\Oh.WAV"
MsgBox "You Walk On Fire !! You Have Red Boots For The Fire !!.", vbOKOnly + vbInformation, "On Fire"
Unload Me
Me.Show
End If
Exit Sub
End If

If E(EN.Text).Tag = "WATER" Then
If WA.Caption = "OFF" Then
s.URL = App.path + "\Data Game\Sounds\Oh.WAV"
MsgBox "You Walk On Water !! You Have Blue Boots For The Water !!.", vbOKOnly + vbInformation, "On Water"
Unload Me
Me.Show
End If
Exit Sub
End If

If E(EN.Text).Tag = "ROCK" Then
Dim BlockTemp As String
If DI.Text = "R" Then
If E(EN.Text + 1).Tag = "FIRE" Then
E(EN.Text).Tag = E(Int(EN.Text) - 1).Tag
E(EN.Text + 1).Tag = E(Int(EN.Text) - 1).Tag
Exit Sub
End If
If E(EN.Text + 1).Tag = "WATER" Then
E(EN.Text).Tag = E(Int(EN.Text) - 1).Tag
E(EN.Text + 1).Tag = E(Int(EN.Text) - 1).Tag
Exit Sub
End If
If E(EN.Text + 1).Tag = "E" Then
BlockTemp = E(EN.Text + 1).ToolTipText
E(EN.Text + 1).Tag = "ROCK"
E(EN.Text + 1).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = BlockTemp
E(Int(EN.Text) + 1).Picture = LoadPicture(App.path + "\Data Game\Item\ROCK" + E(Int(EN.Text) + 1).ToolTipText + ".bmp")
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(Int(EN.Text)).ToolTipText + ".bmp")
End If
End If
If DI.Text = "L" Then
If E(EN.Text - 1).Tag = "FIRE" Then
E(EN.Text).Tag = E(Int(EN.Text) + 1).Tag
E(EN.Text - 1).Tag = E(Int(EN.Text) + 1).Tag
Exit Sub
End If
If E(EN.Text - 1).Tag = "WATER" Then
E(EN.Text).Tag = E(Int(EN.Text) + 1).Tag
E(EN.Text - 1).Tag = E(Int(EN.Text) + 1).Tag
Exit Sub
End If
If E(EN.Text - 1).Tag = "E" Then
BlockTemp = E(EN.Text - 1).ToolTipText
E(EN.Text - 1).Tag = "ROCK"
E(EN.Text - 1).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = BlockTemp
E(Int(EN.Text) - 1).Picture = LoadPicture(App.path + "\Data Game\Item\ROCK" + E(Int(EN.Text) - 1).ToolTipText + ".bmp")
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(Int(EN.Text)).ToolTipText + ".bmp")
End If
End If
If DI.Text = "D" Then
If E(EN.Text + 24).Tag = "FIRE" Then
E(EN.Text).Tag = E(Int(EN.Text) - 24).Tag
E(EN.Text + 24).Tag = E(Int(EN.Text) - 24).Tag
Exit Sub
End If
If E(EN.Text + 24).Tag = "WATER" Then
E(EN.Text).Tag = E(Int(EN.Text) - 24).Tag
E(EN.Text + 24).Tag = E(Int(EN.Text) - 24).Tag
Exit Sub
End If
If E(EN.Text + 24).Tag = "E" Then
BlockTemp = E(EN.Text + 24).ToolTipText
E(EN.Text + 24).Tag = "ROCK"
E(EN.Text + 24).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = BlockTemp
E(Int(EN.Text) + 24).Picture = LoadPicture(App.path + "\Data Game\Item\ROCK" + E(Int(EN.Text) + 24).ToolTipText + ".bmp")
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(Int(EN.Text)).ToolTipText + ".bmp")
End If
End If
If DI.Text = "U" Then
If E(EN.Text - 24).Tag = "FIRE" Then
E(EN.Text).Tag = E(Int(EN.Text) + 24).Tag
E(EN.Text - 24).Tag = E(Int(EN.Text) + 24).Tag
Exit Sub
End If
If E(EN.Text - 24).Tag = "WATER" Then
E(EN.Text).Tag = E(Int(EN.Text) + 24).Tag
E(EN.Text - 24).Tag = E(Int(EN.Text) + 24).Tag
Exit Sub
End If
If E(EN.Text - 24).Tag = "E" Then
BlockTemp = E(EN.Text - 24).ToolTipText
E(EN.Text - 24).Tag = "ROCK"
E(EN.Text - 24).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = BlockTemp
E(Int(EN.Text) - 24).Picture = LoadPicture(App.path + "\Data Game\Item\ROCK" + E(Int(EN.Text) - 24).ToolTipText + ".bmp")
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(Int(EN.Text)).ToolTipText + ".bmp")
End If
End If
If E(EN.Text).Tag = "ROCK" Then
If DI.Text = "L" And E(EN.Text + 1).Tag = "E" Then
BlockTemp = E(EN.Text + 1).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text + 1).Tag = "ROCK"
E(EN.Text + 1).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).ToolTipText = BlockTemp
Exit Sub
End If
If DI.Text = "R" And E(EN.Text - 1).Tag = "E" Then
BlockTemp = E(EN.Text - 1).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text - 1).Tag = "ROCK"
E(EN.Text - 1).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).ToolTipText = BlockTemp
Exit Sub
End If
If DI.Text = "U" And E(EN.Text + 24).Tag = "E" Then
BlockTemp = E(EN.Text + 24).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text + 24).Tag = "ROCK"
E(EN.Text + 24).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).ToolTipText = BlockTemp
Exit Sub
End If
If DI.Text = "D" And E(EN.Text - 24).Tag = "E" Then
BlockTemp = E(EN.Text - 24).ToolTipText
E(EN.Text).Tag = "E"
E(EN.Text - 24).Tag = "ROCK"
E(EN.Text - 24).ToolTipText = E(EN.Text).ToolTipText
E(EN.Text).ToolTipText = BlockTemp
Exit Sub
End If
End If
s.URL = App.path + "\Data Game\Sounds\ROCK.WAV"
End If

If E(EN.Text).Tag = "GETFIRE" Then
FI.Caption = "ON"
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\" + E(Int(EN.Text)).Tag + E(Int(EN.Text)).ToolTipText + ".bmp")
Exit Sub
End If

If E(EN.Text).Tag = "GETWATER" Then
WA.Caption = "ON"
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\" + E(Int(EN.Text)).Tag + E(Int(EN.Text)).ToolTipText + ".bmp")
Exit Sub
End If

If E(EN.Text).Tag = "GETICE" Then
IC.Caption = "ON"
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\" + E(Int(EN.Text)).Tag + E(Int(EN.Text)).ToolTipText + ".bmp")
Exit Sub
End If

If E(EN.Text).Tag = "TIPS" Then
HLP.Caption = E(EN.Text).ToolTipText
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
Exit Sub
End If

If E(EN.Text).Tag = "GETRED" Then
KR.Caption = "ON"
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
Exit Sub
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
End If

If E(EN.Text).Tag = "GETBLUE" Then
KB.Caption = "ON"
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\" + E(Int(EN.Text)).Tag + E(Int(EN.Text)).ToolTipText + ".bmp")
Exit Sub
End If

If E(EN.Text).Tag = "GETYELLOW" Then
KY.Caption = "ON"
E(EN.Text).Tag = "E"
E(EN.Text).ToolTipText = "E"
s.URL = App.path + "\Data Game\Sounds\GET.WAV"
E(Int(EN.Text)).Picture = LoadPicture(App.path + "\Data Game\Item\" + E(Int(EN.Text)).Tag + E(Int(EN.Text)).ToolTipText + ".bmp")
Exit Sub
End If

If E(EN.Text).Tag = "ICER" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
CON.Text = "D"
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICEL" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
CON.Text = "A"
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICEU" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
CON.Text = "W"
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICED" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
CON.Text = "S"
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICE" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "R" Then
CON.Text = "D"
End If
If DI.Text = "L" Then
CON.Text = "A"
End If
If DI.Text = "U" Then
CON.Text = "W"
End If
If DI.Text = "D" Then
CON.Text = "S"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICEUR" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "U" Then
CON.Text = "D"
End If
If DI.Text = "L" Then
CON.Text = "S"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICEUL" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "U" Then
CON.Text = "A"
End If
If DI.Text = "R" Then
CON.Text = "S"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICEDR" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "D" Then
CON.Text = "D"
End If
If DI.Text = "L" Then
CON.Text = "W"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICEDL" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "D" Then
CON.Text = "A"
End If
If DI.Text = "R" Then
CON.Text = "W"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICERU" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "R" Then
CON.Text = "W"
End If
If DI.Text = "D" Then
CON.Text = "A"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICERD" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "R" Then
CON.Text = "S"
End If
If DI.Text = "U" Then
CON.Text = "A"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICELU" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "L" Then
CON.Text = "W"
End If
If DI.Text = "D" Then
CON.Text = "D"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "ICELD" Then
If IC.Caption = "OFF" Then
MsgBox "You Walk On Ice !! You Have White Boots For The Ice !!.", vbOKOnly + vbInformation, "On Ice"
Unload Me
Me.Show
Else
If DI.Text = "L" Then
CON.Text = "S"
End If
If DI.Text = "U" Then
CON.Text = "D"
End If
Exit Sub
End If
End If

If E(EN.Text).Tag = "H" Then
If DI.Text = "R" Then
CON.Text = "A"
End If
If DI.Text = "L" Then
CON.Text = "D"
End If
If DI.Text = "U" Then
CON.Text = "S"
End If
If DI.Text = "D" Then
CON.Text = "W"
End If
Exit Sub
End If

If E(EN.Text).Tag = "H2" Then
If DI.Text = "R" Then
y.Left = y.Left - E(0).Width
EN.Text = EN.Text - 1
L.Left = L.Left + E(0).Width
End If
If DI.Text = "L" Then
CON.Text = "D"
End If
If DI.Text = "U" Then
y.Top = y.Top + E(0).Width
EN.Text = EN.Text + 24
L.Top = L.Top - E(0).Width
End If
If DI.Text = "D" Then
CON.Text = "W"
End If
Exit Sub
End If

If E(EN.Text).Tag = "LU" Then
For i = 0 To 431
If E(i).Tag = "H2" Then
E(i).Tag = "E2"
GoTo 5
End If
If E(i).Tag = "E2" Then
E(i).Tag = "H2"
End If
5:
Next
If DI.Text = "L" Then
If E(EN.Text - 1).Tag = "E" Then
CON.Text = "A"
Else
CON.Text = "D"
Exit Sub
End If
End If
If DI.Text = "R" Then
If E(EN.Text + 1).Tag = "E" Then
CON.Text = "D"
Else
CON.Text = "A"
Exit Sub
End If
End If
If DI.Text = "U" Then
If E(EN.Text - 24).Tag = "E" Then
CON.Text = "W"
Else
CON.Text = "S"
Exit Sub
End If
End If
If DI.Text = "D" Then
If E(EN.Text + 24).Tag = "E" Then
CON.Text = "S"
Else
CON.Text = "W"
Exit Sub
End If
End If
Exit Sub
End If
End Sub

Private Sub DAROI_Timer()
On Error Resume Next

y.Left = E(EN.Text).Left
y.Top = E(EN.Text).Top
L.Left = -y.Left + E(0).Width * PlayerPos
L.Top = -y.Top + E(0).Height * PlayerPos

DAROI.Interval = 100
For i = 0 To 431

DoEvents
Select Case E(i).Tag
Case Is = "E"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\E" + E(i).ToolTipText + ".bmp")
Case Is = "DH"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\DH" + E(i).ToolTipText + ".bmp")
Case Is = "END"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\END" + E(i).ToolTipText + ".bmp")
Case Is = "GETFIRE"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\GETFIRE" + E(i).ToolTipText + ".bmp")
Case Is = "LU"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\LU" + E(i).ToolTipText + ".bmp")
Case Is = "H2"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\H2" + E(i).ToolTipText + ".bmp")
Case Is = "E2"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\E2" + E(i).ToolTipText + ".bmp")
Case Is = "GETWATER"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\GETWATER" + E(i).ToolTipText + ".bmp")
Case Is = "GETICE"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\GETICE" + E(i).ToolTipText + ".bmp")
Case Is = "FIRE"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\FIRE" + E(i).ToolTipText + ".bmp")
Case Is = "TIPS"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\TIPS.bmp")
Case Is = "ICE"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICE" + E(i).ToolTipText + ".bmp")
Case Is = "H"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\BLOCK" + E(i).ToolTipText + ".bmp")
Case Is = "ROCK"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ROCK" + E(i).ToolTipText + ".bmp")
Case Is = "WATER"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\WATER" + E(i).ToolTipText + ".bmp")
Case Is = "ICEU"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICEU" + E(i).ToolTipText + ".bmp")
Case Is = "ICED"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICED" + E(i).ToolTipText + ".bmp")
Case Is = "ICEL"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICEL" + E(i).ToolTipText + ".bmp")
Case Is = "ICER"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICER" + E(i).ToolTipText + ".bmp")
Case Is = "ICEUR"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICEUR" + E(i).ToolTipText + ".bmp")
Case Is = "ICEUL"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICEUL" + E(i).ToolTipText + ".bmp")
Case Is = "ICEDR"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICEDR" + E(i).ToolTipText + ".bmp")
Case Is = "ICEDL"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICEDL" + E(i).ToolTipText + ".bmp")
Case Is = "ICERU"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICERU" + E(i).ToolTipText + ".bmp")
Case Is = "ICERD"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICERD" + E(i).ToolTipText + ".bmp")
Case Is = "ICELU"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICELU" + E(i).ToolTipText + ".bmp")
Case Is = "ICELD"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\ICELD" + E(i).ToolTipText + ".bmp")
Case Is = "GETRED"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\GR" + E(i).ToolTipText + ".bmp")
Case Is = "GETBLUE"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\GB" + E(i).ToolTipText + ".bmp")
Case Is = "GETYELLOW"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\GY" + E(i).ToolTipText + ".bmp")
Case Is = "DOORHR"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\HKR.bmp")
Case Is = "DOORHB"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\HKB.bmp")
Case Is = "DOORHY"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\HKY.bmp")
Case Is = "COM"
E(i).Picture = LoadPicture(App.path + "\Data Game\Item\COM" + E(i).ToolTipText + ".bmp")
End Select

Next
End Sub

Private Sub DHB_Change()
DHH.Caption = DHB.Caption + " \ " + ALLDHB.Caption
End Sub


Private Sub E_Click(Index As Integer)
L.SetFocus
End Sub

Private Sub FFAE_Click()
End
End Sub

Private Sub FI_Change()
If FI.Caption = "ON" Then
A.Picture = LoadPicture(App.path + "\Data Game\Item\GETFIREe.bmp")
Else
A.Picture = DEL.Picture
End If
End Sub

Private Sub Form_Load()
On Error GoTo 6

SkipStory = False

NUM = FreeFile
Open App.path + "\Data Game\Save.ZAN" For Input As NUM
Dim x As String
Input #NUM, x
LV.Caption = x
Close NUM

ChangeSizeMap 1
PlayerPos = 3
BackStory.Picture = LoadPicture(App.path + "\Data Game\Story\Back.jpg")

Me.Show

For i = 1 To StoryNum \ 2
Input #99, x
NameTalk.Caption = x + " :"
WhoTalk.Picture = LoadPicture(App.path + "\Data Game\Story\" + x + ".emf")
Input #99, x
TextTalking.Caption = x
NextStory = False
Do While NextStory = False And SkipStory = False
' wait to skip
DoEvents
Loop
Next

SkipStory = True

Picture7.Visible = False

6:
Close #99
If EndGame = True Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo 5
NUM = FreeFile
Open App.path + "\Data Game\Save.ZAN" For Output As NUM
Write #NUM, LV.Caption
Close NUM
5:
End Sub

Private Sub GSFS_Click()
Form2.Show
End Sub

Private Sub HLP_Change()
HL.Caption = HLP.Caption
End Sub

Private Sub IC_Change()
If IC.Caption = "ON" Then
C.Picture = LoadPicture(App.path + "\Data Game\Item\GETICEe.bmp")
Else
C.Picture = DEL.Picture
End If
End Sub

Private Sub KB_Change()
If KB.Caption = "ON" Then
B1.Picture = LoadPicture(App.path + "\Data Game\Item\GBe.bmp")
Else
B1.Picture = DEL.Picture
End If
End Sub

Private Sub KR_Change()
If KR.Caption = "ON" Then
A1.Picture = LoadPicture(App.path + "\Data Game\Item\GRe.bmp")
Else
A1.Picture = DEL.Picture
End If
End Sub

Private Sub KY_Change()
If KY.Caption = "ON" Then
C1.Picture = LoadPicture(App.path + "\Data Game\Item\GYe.bmp")
Else
C1.Picture = DEL.Picture
End If
End Sub

Private Sub L_KeyPress(KeyAscii As Integer)
If PressWait < 2 Then
PressWait = PressWait + 1
Exit Sub
Else
PressWait = 0
End If

Select Case KeyAscii
Case Is = 100: CON.Text = "D"
Case Is = 97: CON.Text = "A"
Case Is = 119: CON.Text = "W"
Case Is = 115: CON.Text = "S"
Case Is = 101: CON.Text = "E"
Case Is = 113: CON.Text = "Q"
End Select
End Sub

Private Sub L_KeyUp(KeyCode As Integer, Shift As Integer)
PressWait = 9
End Sub

Private Sub LV_Change()
On Error GoTo 5

If SkipStory = True Then Exit Sub

Open App.path + "\Data Game\Levels\" + LV.Caption + ".Zan" For Input As #99
Dim x As String
Input #99, x
PW.Caption = x
Input #99, x
ALLDHB.Caption = x
For i = 0 To 431
Input #99, x
E(i).Tag = x
Input #99, x
E(i).ToolTipText = x
E(i).DataField = E(i).ToolTipText
If E(i).Tag = "DOORHB" Then E(i).ToolTipText = "LOCKED"
If E(i).Tag = "DOORHR" Then E(i).ToolTipText = "LOCKED"
If E(i).Tag = "DOORHY" Then E(i).ToolTipText = "LOCKED"
Next
Input #99, x
Initialize_Music (25)
Load_Music (Int(x))
PlayMusic
SetMusic (100)
Input #99, x
StoryNum = Int(x)

For i = 0 To 431
If E(i).Tag = "COM" Then
If E(i).ToolTipText = "R" Then E(i).DataField = E(i + 1).DataField
If E(i).ToolTipText = "L" Then E(i).DataField = E(i - 1).DataField
If E(i).ToolTipText = "D" Then E(i).DataField = E(i + 24).DataField
If E(i).ToolTipText = "U" Then E(i).DataField = E(i - 24).DataField
End If
Next

LVL.Caption = LV.Caption
GoTo 6
5:
DAROI.Interval = 0
DAON.Interval = 0
SkipStory = True
Form4.Show
EndGame = True
6:
End Sub

Private Sub mnuAbout_Click()
Form5.Show
End Sub

Private Sub mnuCheats_Click()
CH.Visible = True
End Sub

Private Sub mnuLarge_Click()
ChangeSizeMap 2
PlayerPos = 1
End Sub

Private Sub mnuMuteMus_Click()
SetMusic (0)
End Sub

Private Sub mnuNewGame_Click()
Unload Me
Me.Show
End Sub

Private Sub mnuNormal_Click()
ChangeSizeMap 1
PlayerPos = 3
End Sub

Private Sub mnuSave_Click()
On Error GoTo 5
NUM = FreeFile
Open App.path + "\Data Game\Save.ZAN" For Output As NUM
Write #NUM, LV.Caption
Close NUM
5:
End Sub

Private Sub mnuSmall_Click()
ChangeSizeMap 0.5
PlayerPos = 6
End Sub

Private Sub mnuVol1_Click()
SetMusic (50)
End Sub

Private Sub mnuVol2_Click()
SetMusic (100)
End Sub

Private Sub mnuVol3_Click()
SetMusic (200)
End Sub

Private Sub mnuVolOff_Click()
s.settings.mute = True
End Sub

Private Sub mnuVolOn_Click()
s.settings.mute = False
End Sub

Private Sub Timer1_Timer()
If IsMusicAtEnd = True Then PlayMusic
End Sub

Private Sub WA_Change()
If WA.Caption = "ON" Then
B.Picture = LoadPicture(App.path + "\Data Game\Item\GETWATERe.bmp")
Else
B.Picture = DEL.Picture
End If
End Sub

Sub ChangeSizeMap(Size As Single)
Dim RowNum As Integer
RowNum = 24

Dim i, ObjTop, ObjLeft As Integer

Dim ObjNum As Integer
For i = 0 To E.Count - 1
If E(i).Visible = True Then ObjNum = ObjNum + 1
E(i).Stretch = True
E(i).Width = 495
E(i).Height = 495
Next

E(0).Width = E(0).Width * Size
E(0).Height = E(0).Height * Size

ObjTop = E(0).Top

For i = 1 To ObjNum - 1

E(i).Width = E(0).Width
E(i).Height = E(0).Height

ObjLeft = ObjLeft + 1
If ObjLeft = RowNum Then
ObjLeft = 0
ObjTop = ObjTop + (E(0).Height)
End If

E(i).Left = ObjLeft * E(i).Width
E(i).Top = ObjTop

Next

y.Width = E(0).Width
y.Height = E(0).Height

L.Height = ObjTop + E(0).Height
L.Width = E(0).Width * RowNum
End Sub
