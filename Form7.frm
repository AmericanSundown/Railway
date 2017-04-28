VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellation "
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8760
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Visible = False
Dim Beginpage, EndPage, NumCopies, orientation, i
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.ShowPrinter
Beginpage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
orientation = CommonDialog1.orientation
For i = 1 To NumCopies
Form7.PrintForm
Next
Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub Form_Load()
Text6.Text = n7
Text3.Text = Temp5
Text5.Text = Temp5
End Sub

Private Sub Text6_Change()
Label9.Caption = "Rs." & Val(Text6.Text) * 20
End Sub
