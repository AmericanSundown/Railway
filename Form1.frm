VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Railway Reservation System - Login"
   ClientHeight    =   2550
   ClientLeft      =   6165
   ClientTop       =   4740
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5640
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text3.Text = "Error" And Text4.Text = "404" Then
Unload Me
Form14.Show
Else
MsgBox ("Invalid Username/Password")
End If
End Sub

Private Sub Command2_Click()
End
End Sub


