VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Railway Reservation System"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form12"
   Picture         =   "Form122.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Railway Reservation System"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command2_Click()
MDIForm1.Show
End Sub

Private Sub Command3_Click()
End
End Sub

