VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   1770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1230
      Top             =   480
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   525
      Left            =   270
      TabIndex        =   11
      Top             =   1200
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   660
      Left            =   720
      ScaleHeight     =   600
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   300
      Width           =   345
      Begin VB.Label lblnum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   9
         Left            =   180
         TabIndex        =   10
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   8
         Left            =   1020
         TabIndex        =   9
         Top             =   1410
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   7
         Left            =   600
         TabIndex        =   8
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   5
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   4
         Left            =   540
         TabIndex        =   4
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   2
         Left            =   900
         TabIndex        =   3
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   1
         Left            =   450
         TabIndex        =   2
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblnum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   315
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Reset()

  Dim i As Long
  
  lblnum(0).Top = 0
  lblnum(0).Left = 0
  
  For i = 1 To 9
    lblnum(i).Top = lblnum(i - 1).Top + lblnum(i - 1).Height
    lblnum(i).Left = 0
  Next
  
End Sub

Private Sub cmdExit_Click()
  
  Timer1.Enabled = False
  Unload Me

End Sub


Private Sub cmdReset_Click()
  Reset
End Sub

Private Sub cmdRun_Click()

  Timer1.Enabled = True
  
End Sub


Private Sub Form_Load()

  Reset
  
  Timer1.Enabled = True
  
End Sub


Private Sub Timer1_Timer()

  Dim i As Integer
  
  For i = 0 To 9
    lblnum(i).Top = lblnum(i).Top - 10
  Next
  
 'And this line make it look seamless...
  If lblnum(9).Top <= 0 Then lblnum(0).Top = lblnum(9).Top + lblnum(9).Height
  
  If lblnum(9).Top + lblnum(9).Height < 0 Then Reset
  
End Sub


