VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Railway Reservation System"
   ClientHeight    =   3255
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6360
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2246.659
   ScaleMode       =   0  'User
   ScaleWidth      =   5972.369
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1290
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RAILWAY RESERVATION SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   638
      TabIndex        =   2
      Top             =   240
      Width           =   5085
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1822.175
      Y2              =   1822.175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

