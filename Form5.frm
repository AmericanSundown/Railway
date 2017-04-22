VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4725
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "Passenger_Name"
         Caption         =   "Passenger Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Age"
         Caption         =   "Age"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Sex"
         Caption         =   "Sex"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Senior_Citizen"
         Caption         =   "Senior Citizen"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "seat_no"
         Caption         =   "Seat No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Passenger_Name"
         Caption         =   "Passenger_Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Age"
         Caption         =   "Age"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Sex"
         Caption         =   "Sex"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Senior_Citizen"
         Caption         =   "Senior_Citizen"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   0
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape77 
      Height          =   135
      Left            =   120
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   135
   End
   Begin VB.Shape Shape72 
      Height          =   135
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   135
   End
   Begin VB.Line Line15 
      X1              =   0
      X2              =   360
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line8 
      X1              =   8400
      X2              =   8760
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line7 
      X1              =   8400
      X2              =   8760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line6 
      X1              =   8760
      X2              =   8760
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   360
      Y1              =   5880
      Y2              =   5880
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim cn1 As New ADODB.Connection
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
Form5.PrintForm
Next
Exit Sub
ErrHandler:
Exit Sub
End Sub


Private Sub Form_Load()
Label9.Caption = "Rs." & "" & temp3
Text3.Text = Temp2
s = " select * from reservation where PNR_NO = " & Text3.Text & " "
connect (s)
Set Text1.DataSource = rs
Text1.DataField = "Train_No"
Set Text2.DataSource = rs
Text2.DataField = "Date_travel"
Set Text3.DataSource = rs
Text3.DataField = "PNR_NO"
Set Text4.DataSource = rs
Text4.DataField = "Class"
Set Text18.DataSource = rs
Text18.DataField = "Train_Name"
Set Text19.DataSource = rs
Text19.DataField = "From"
Set Text20.DataSource = rs
Text20.DataField = "To"
Set Label22.DataSource = rs
Label22.DataField = "date_travel"
Set DataGrid1.DataSource = rs
SQL = "select * from timings where train_no = " & Text1.Text & ""
Set cn1 = New ADODB.Connection
cn1.CursorLocation = adUseClient
cn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
cn1.Open
Set rs1 = New ADODB.Recordset
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
rs1.ActiveConnection = cn1
rs1.Open SQL
Set Label17.DataSource = rs1
Label17.DataField = "Distance"
Set Label20.DataSource = rs1
Label20.DataField = "Arrival_Time"
Set Label24.DataSource = rs1
Label24.DataField = "departure_time"

End Sub


