VERSION 5.00
Begin VB.Form frmAge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Moon's Age"
   ClientHeight    =   5616
   ClientLeft      =   996
   ClientTop       =   2316
   ClientWidth     =   6432
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5616
   ScaleWidth      =   6432
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   492
      Left            =   4560
      TabIndex        =   18
      Top             =   4200
      Width           =   1452
   End
   Begin VB.CommandButton cmdComp 
      Caption         =   "Recompute"
      Height          =   492
      Left            =   2400
      TabIndex        =   17
      Top             =   4200
      Width           =   1452
   End
   Begin VB.OptionButton optAM 
      Caption         =   "PM"
      Height          =   252
      Index           =   1
      Left            =   5520
      TabIndex        =   10
      Top             =   1560
      Width           =   612
   End
   Begin VB.OptionButton optAM 
      Caption         =   "AM"
      Height          =   252
      Index           =   0
      Left            =   5520
      TabIndex        =   9
      Top             =   1200
      Value           =   -1  'True
      Width           =   612
   End
   Begin VB.TextBox txtMin 
      Height          =   288
      Left            =   4560
      TabIndex        =   8
      Top             =   1440
      Width           =   612
   End
   Begin VB.TextBox txtHour 
      Height          =   288
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   612
   End
   Begin VB.TextBox txtYear 
      Height          =   288
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   612
   End
   Begin VB.TextBox txtDay 
      Height          =   288
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   492
   End
   Begin VB.Frame Frame1 
      Caption         =   "Month"
      Height          =   4812
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1572
      Begin VB.OptionButton optMonth 
         Caption         =   "December"
         Height          =   252
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   4320
         Width           =   1332
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "November"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   3960
         Width           =   1332
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "October"
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   3600
         Width           =   1212
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "September"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   1332
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "August"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   1212
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "July"
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   972
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "June"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   972
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "May"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   972
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "April"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   972
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "March"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   972
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "February"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1332
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "January"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1332
      End
   End
   Begin VB.Label lblLoc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vancouver, British Columbia"
      Height          =   612
      Left            =   2040
      TabIndex        =   31
      Top             =   480
      Width           =   4212
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "hours"
      Height          =   252
      Left            =   5400
      TabIndex        =   16
      Top             =   3000
      Width           =   612
   End
   Begin VB.Label lblHours 
      Height          =   252
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   492
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "days"
      Height          =   252
      Left            =   4080
      TabIndex        =   14
      Top             =   3000
      Width           =   492
   End
   Begin VB.Label lblDays 
      Height          =   252
      Left            =   3720
      TabIndex        =   13
      Top             =   3000
      Width           =   252
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Moon's age    ="
      Height          =   252
      Left            =   2280
      TabIndex        =   12
      Top             =   3000
      Width           =   1332
   End
   Begin VB.Label lblDOW 
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      Height          =   252
      Left            =   3240
      TabIndex        =   11
      Top             =   2160
      Width           =   1572
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      Height          =   252
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hour"
      Height          =   252
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   612
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   252
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
      Height          =   252
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   612
   End
End
Attribute VB_Name = "frmAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Hide
End Sub

Private Sub cmdComp_Click()
    Dim i As Integer
    Dim j As Integer
    Dim x As Double
    Dim y As Double
    Dim ADYear As Integer
    Dim ADMonth As Integer
    Dim ADday As Integer
    Dim ADhour As Integer
    Dim ADminute As Integer
    For i = 0 To 11
        If optMonth(i).Value Then
            ADMonth = i + 1
            Exit For
        End If
    Next i
    i = RangeCheck(txtYear, 640, 4503)
    If i <= 0 Then
        Exit Sub
    End If
    ADYear = i
    If ADMonth = 2 Then
        j = 28 + IsLeap(ADYear)
    Else
        j = ndmnth(ADMonth - 1)
    End If
    i = RangeCheck(txtDay, 1, j)
    If i <= 0 Then
        Exit Sub
    End If
    ADday = i
    i = RangeCheck(txtHour, 0, 12)
    If i < 0 Then
        Exit Sub
    End If
    ADhour = i
    i = RangeCheck(txtMin, 0, 59)
    If i < 0 Then   'SKA 2/20/95: changed from i<=0
        Exit Sub
    End If
    ADminute = i
    If optAM(0).Value Then
        ampm = 0
    Else
        ampm = 1
        If ADhour <> 12 Then
            ADhour = ADhour + 12
        End If
    End If
    x = Age(ADYear, ADMonth, ADday, ADhour, ADminute, dayOfWeek)
    y = Fix(x)
    lblDays.Caption = y
    lblHours.Caption = Format((x - y) * 24#, "#0.0")
    lblDOW.Caption = weekdayName(dayOfWeek)
End Sub

Private Sub Form_Load()
    top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub


