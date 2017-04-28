VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "s"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14730
   LinkTopic       =   "Form14"
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   13
      Top             =   8160
      Width           =   5295
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   10320
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   10320
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   10320
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hello Admin"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
