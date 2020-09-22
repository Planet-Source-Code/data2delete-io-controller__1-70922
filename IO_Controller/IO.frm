VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IO_Controller"
   ClientHeight    =   7275
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11220
   Icon            =   "IO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   2625
      Top             =   6405
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Blink Line 1"
      Height          =   645
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   6405
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      Height          =   645
      Left            =   8715
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   6405
      Width           =   2115
   End
   Begin VB.CheckBox chk8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   5670
      TabIndex        =   89
      Top             =   4935
      Width           =   960
   End
   Begin VB.CheckBox chk7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   5670
      TabIndex        =   88
      Top             =   3465
      Width           =   960
   End
   Begin VB.CheckBox chk6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   5670
      TabIndex        =   87
      Top             =   1995
      Width           =   960
   End
   Begin VB.CheckBox chk5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   5670
      TabIndex        =   86
      Top             =   525
      Width           =   960
   End
   Begin VB.CheckBox chk4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   210
      TabIndex        =   85
      Top             =   4935
      Width           =   960
   End
   Begin VB.CheckBox chk3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   210
      TabIndex        =   84
      Top             =   3465
      Width           =   960
   End
   Begin VB.CheckBox chk2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   210
      TabIndex        =   83
      Top             =   1995
      Width           =   960
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timer Off"
      Height          =   540
      Left            =   210
      TabIndex        =   82
      Top             =   525
      Width           =   960
   End
   Begin VB.Timer tmr8 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   6720
      Top             =   5460
   End
   Begin VB.Timer tmr4 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   1155
      Top             =   5460
   End
   Begin VB.Timer tmr7 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   6720
      Top             =   3990
   End
   Begin VB.Timer tmr3 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   1155
      Top             =   3990
   End
   Begin VB.Timer tmr6 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   6720
      Top             =   2520
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   1155
      Top             =   2520
   End
   Begin VB.Timer tmr5 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   6720
      Top             =   1050
   End
   Begin VB.ComboBox Combo48 
      Height          =   315
      Left            =   7455
      TabIndex        =   79
      Text            =   "12"
      Top             =   5565
      Width           =   1065
   End
   Begin VB.ComboBox Combo47 
      Height          =   315
      Left            =   8610
      TabIndex        =   78
      Text            =   "00"
      Top             =   5565
      Width           =   1065
   End
   Begin VB.ComboBox Combo46 
      Height          =   315
      Left            =   9765
      TabIndex        =   77
      Text            =   "AM"
      Top             =   5565
      Width           =   1065
   End
   Begin VB.ComboBox Combo45 
      Height          =   315
      Left            =   7455
      TabIndex        =   76
      Text            =   "12"
      Top             =   5880
      Width           =   1065
   End
   Begin VB.ComboBox Combo44 
      Height          =   315
      Left            =   8610
      TabIndex        =   75
      Text            =   "00"
      Top             =   5880
      Width           =   1065
   End
   Begin VB.ComboBox Combo43 
      Height          =   315
      Left            =   9765
      TabIndex        =   74
      Text            =   "PM"
      Top             =   5880
      Width           =   1065
   End
   Begin VB.ComboBox Combo42 
      Height          =   315
      Left            =   1890
      TabIndex        =   71
      Text            =   "12"
      Top             =   5565
      Width           =   1065
   End
   Begin VB.ComboBox Combo41 
      Height          =   315
      Left            =   3045
      TabIndex        =   70
      Text            =   "00"
      Top             =   5565
      Width           =   1065
   End
   Begin VB.ComboBox Combo40 
      Height          =   315
      Left            =   4200
      TabIndex        =   69
      Text            =   "AM"
      Top             =   5565
      Width           =   1065
   End
   Begin VB.ComboBox Combo39 
      Height          =   315
      Left            =   1890
      TabIndex        =   68
      Text            =   "12"
      Top             =   5880
      Width           =   1065
   End
   Begin VB.ComboBox Combo38 
      Height          =   315
      Left            =   3045
      TabIndex        =   67
      Text            =   "00"
      Top             =   5880
      Width           =   1065
   End
   Begin VB.ComboBox Combo37 
      Height          =   315
      Left            =   4200
      TabIndex        =   66
      Text            =   "PM"
      Top             =   5880
      Width           =   1065
   End
   Begin VB.ComboBox Combo36 
      Height          =   315
      Left            =   7455
      TabIndex        =   63
      Text            =   "12"
      Top             =   4095
      Width           =   1065
   End
   Begin VB.ComboBox Combo35 
      Height          =   315
      Left            =   8610
      TabIndex        =   62
      Text            =   "00"
      Top             =   4095
      Width           =   1065
   End
   Begin VB.ComboBox Combo34 
      Height          =   315
      Left            =   9765
      TabIndex        =   61
      Text            =   "AM"
      Top             =   4095
      Width           =   1065
   End
   Begin VB.ComboBox Combo33 
      Height          =   315
      Left            =   7455
      TabIndex        =   60
      Text            =   "12"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.ComboBox Combo32 
      Height          =   315
      Left            =   8610
      TabIndex        =   59
      Text            =   "00"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.ComboBox Combo31 
      Height          =   315
      Left            =   9765
      TabIndex        =   58
      Text            =   "PM"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.ComboBox Combo30 
      Height          =   315
      Left            =   1890
      TabIndex        =   55
      Text            =   "12"
      Top             =   4095
      Width           =   1065
   End
   Begin VB.ComboBox Combo29 
      Height          =   315
      Left            =   3045
      TabIndex        =   54
      Text            =   "00"
      Top             =   4095
      Width           =   1065
   End
   Begin VB.ComboBox Combo28 
      Height          =   315
      Left            =   4200
      TabIndex        =   53
      Text            =   "AM"
      Top             =   4095
      Width           =   1065
   End
   Begin VB.ComboBox Combo27 
      Height          =   315
      Left            =   1890
      TabIndex        =   52
      Text            =   "12"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.ComboBox Combo26 
      Height          =   315
      Left            =   3045
      TabIndex        =   51
      Text            =   "00"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.ComboBox Combo25 
      Height          =   315
      Left            =   4200
      TabIndex        =   50
      Text            =   "PM"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.ComboBox Combo24 
      Height          =   315
      Left            =   7455
      TabIndex        =   47
      Text            =   "12"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.ComboBox Combo23 
      Height          =   315
      Left            =   8610
      TabIndex        =   46
      Text            =   "00"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.ComboBox Combo22 
      Height          =   315
      Left            =   9765
      TabIndex        =   45
      Text            =   "AM"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.ComboBox Combo21 
      Height          =   315
      Left            =   7455
      TabIndex        =   44
      Text            =   "12"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.ComboBox Combo20 
      Height          =   315
      Left            =   8610
      TabIndex        =   43
      Text            =   "00"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.ComboBox Combo19 
      Height          =   315
      Left            =   9765
      TabIndex        =   42
      Text            =   "PM"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.ComboBox Combo18 
      Height          =   315
      Left            =   1890
      TabIndex        =   39
      Text            =   "12"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.ComboBox Combo17 
      Height          =   315
      Left            =   3045
      TabIndex        =   38
      Text            =   "00"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.ComboBox Combo16 
      Height          =   315
      Left            =   4200
      TabIndex        =   37
      Text            =   "AM"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.ComboBox Combo15 
      Height          =   315
      Left            =   1890
      TabIndex        =   36
      Text            =   "12"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.ComboBox Combo14 
      Height          =   315
      Left            =   3045
      TabIndex        =   35
      Text            =   "00"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.ComboBox Combo13 
      Height          =   315
      Left            =   4200
      TabIndex        =   34
      Text            =   "PM"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      Left            =   7455
      TabIndex        =   31
      Text            =   "12"
      Top             =   1155
      Width           =   1065
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   8610
      TabIndex        =   30
      Text            =   "00"
      Top             =   1155
      Width           =   1065
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   9765
      TabIndex        =   29
      Text            =   "AM"
      Top             =   1155
      Width           =   1065
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   7455
      TabIndex        =   28
      Text            =   "12"
      Top             =   1470
      Width           =   1065
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   8610
      TabIndex        =   27
      Text            =   "00"
      Top             =   1470
      Width           =   1065
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   9765
      TabIndex        =   26
      Text            =   "PM"
      Top             =   1470
      Width           =   1065
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   975
      Left            =   1155
      Top             =   1050
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   4200
      TabIndex        =   24
      Text            =   "PM"
      Top             =   1470
      Width           =   1065
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   3045
      TabIndex        =   23
      Text            =   "00"
      Top             =   1470
      Width           =   1065
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1890
      TabIndex        =   22
      Text            =   "12"
      Top             =   1470
      Width           =   1065
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4200
      TabIndex        =   20
      Text            =   "AM"
      Top             =   1155
      Width           =   1065
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3045
      TabIndex        =   19
      Text            =   "00"
      Top             =   1155
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1890
      TabIndex        =   18
      Text            =   "12"
      Top             =   1155
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   975
      Left            =   5460
      Top             =   0
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8820
      TabIndex        =   16
      Text            =   "Lock2"
      Top             =   4830
      Width           =   2010
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8820
      TabIndex        =   15
      Text            =   "Lock1"
      Top             =   3360
      Width           =   2010
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8820
      TabIndex        =   14
      Text            =   "Coffee Pot"
      Top             =   1890
      Width           =   2010
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8820
      TabIndex        =   13
      Text            =   "Air"
      Top             =   420
      Width           =   2010
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3255
      TabIndex        =   12
      Text            =   "Light3"
      Top             =   4830
      Width           =   2010
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3255
      TabIndex        =   11
      Text            =   "Light2"
      Top             =   3360
      Width           =   2010
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3255
      TabIndex        =   10
      Text            =   "Light1"
      Top             =   1890
      Width           =   2010
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3255
      TabIndex        =   9
      Text            =   "Fan"
      Top             =   420
      Width           =   2010
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Turn Line 8 ON"
      Height          =   645
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4830
      Width           =   2115
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Turn Line 7 ON"
      Height          =   645
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   2115
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Turn Line 6 ON "
      Height          =   645
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1890
      Width           =   2115
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Turn Line 5 ON"
      Height          =   645
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   420
      Width           =   2115
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Turn Line 4 ON"
      Height          =   645
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4830
      Width           =   2115
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Turn Line 3 ON"
      Height          =   645
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   2115
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Turn Line 2 ON"
      Height          =   645
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1890
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "All Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6405
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Turn Line 1 ON"
      Height          =   645
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   420
      Width           =   2115
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   330
      Left            =   630
      TabIndex        =   91
      Top             =   7875
      Width           =   2220
   End
   Begin VB.Line Line9 
      X1              =   4200
      X2              =   4200
      Y1              =   6300
      Y2              =   7140
   End
   Begin VB.Line Line8 
      X1              =   5355
      X2              =   4200
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line Line7 
      X1              =   6510
      X2              =   6510
      Y1              =   6300
      Y2              =   7140
   End
   Begin VB.Line Line6 
      X1              =   5355
      X2              =   6510
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line Line5 
      X1              =   5355
      X2              =   5355
      Y1              =   315
      Y2              =   6300
   End
   Begin VB.Line Line4 
      X1              =   10920
      X2              =   10920
      Y1              =   315
      Y2              =   7140
   End
   Begin VB.Line Line3 
      X1              =   105
      X2              =   10920
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line2 
      X1              =   105
      X2              =   10920
      Y1              =   7140
      Y2              =   7140
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   105
      Y1              =   315
      Y2              =   7140
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6990
      TabIndex        =   81
      Top             =   5670
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6930
      TabIndex        =   80
      Top             =   5985
      Width           =   435
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1425
      TabIndex        =   73
      Top             =   5670
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1365
      TabIndex        =   72
      Top             =   5985
      Width           =   435
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6990
      TabIndex        =   65
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6930
      TabIndex        =   64
      Top             =   4515
      Width           =   435
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1425
      TabIndex        =   57
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1365
      TabIndex        =   56
      Top             =   4515
      Width           =   435
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6990
      TabIndex        =   49
      Top             =   2730
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6930
      TabIndex        =   48
      Top             =   3045
      Width           =   435
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1425
      TabIndex        =   41
      Top             =   2730
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1365
      TabIndex        =   40
      Top             =   3045
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6990
      TabIndex        =   33
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6930
      TabIndex        =   32
      Top             =   1575
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF : "
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1365
      TabIndex        =   25
      Top             =   1575
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ON : "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1425
      TabIndex        =   21
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4410
      TabIndex        =   17
      Top             =   0
      Width           =   585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub PortOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Byte)
Private Declare Sub PortWordOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Integer)
Private Declare Sub PortDWordOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Long)
Private Declare Function PortIn Lib "IO.DLL" (ByVal Port As Integer) As Byte
Private Declare Function PortWordIn Lib "IO.DLL" (ByVal Port As Integer) As Integer
Private Declare Function PortDWordIn Lib "IO.DLL" (ByVal Port As Integer) As Long
Private Declare Sub SetPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte)
Private Declare Sub ClrPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte)
Private Declare Sub NotPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte)
Private Declare Function GetPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte) As Boolean
Private Declare Function RightPortShift Lib "IO.DLL" (ByVal Port As Integer, ByVal Val As Boolean) As Boolean
Private Declare Function LeftPortShift Lib "IO.DLL" (ByVal Port As Integer, ByVal Val As Boolean) As Boolean
Private Declare Function IsDriverInstalled Lib "IO.DLL" () As Boolean

Dim A As Integer
Dim B As Integer
Dim y As Integer

Private Sub chk1_Click()
If chk1.Value = Checked Then
chk1.Caption = "Timer ON"
tmr1.Enabled = True
ElseIf chk1.Value = Unchecked Then
chk1.Caption = "Timer Off"
tmr1.Enabled = False
End If
End Sub

Private Sub chk2_Click()
If chk2.Value = Checked Then
chk2.Caption = "Timer ON"
tmr2.Enabled = True
ElseIf chk2.Value = Unchecked Then
chk2.Caption = "Timer Off"
tmr2.Enabled = False
End If
End Sub

Private Sub chk3_Click()
If chk3.Value = Checked Then
chk3.Caption = "Timer ON"
tmr3.Enabled = True
ElseIf chk3.Value = Unchecked Then
chk3.Caption = "Timer Off"
tmr3.Enabled = False
End If
End Sub

Private Sub chk4_Click()
If chk4.Value = Checked Then
chk4.Caption = "Timer ON"
tmr4.Enabled = True
ElseIf chk4.Value = Unchecked Then
chk4.Caption = "Timer Off"
tmr4.Enabled = False
End If
End Sub

Private Sub chk5_Click()
If chk5.Value = Checked Then
chk5.Caption = "Timer ON"
tmr5.Enabled = True
ElseIf chk5.Value = Unchecked Then
chk5.Caption = "Timer Off"
tmr5.Enabled = False
End If
End Sub

Private Sub chk6_Click()
If chk6.Value = Checked Then
chk6.Caption = "Timer ON"
tmr6.Enabled = True
ElseIf chk6.Value = Unchecked Then
chk6.Caption = "Timer Off"
tmr6.Enabled = False
End If
End Sub

Private Sub chk7_Click()
If chk7.Value = Checked Then
chk7.Caption = "Timer ON"
tmr7.Enabled = True
ElseIf chk7.Value = Unchecked Then
chk7.Caption = "Timer Off"
tmr7.Enabled = False
End If
End Sub

Private Sub chk8_Click()
If chk8.Value = Checked Then
chk8.Caption = "Timer ON"
tmr8.Enabled = True
ElseIf chk8.Value = Unchecked Then
chk8.Caption = "Timer Off"
tmr8.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Command1.BackColor = &H8000000F Then
Command1.BackColor = &HC0C0FF
A = A + 1
Call PortOut(888, A)
Command1.Caption = "Turn Line 1 Off"
ElseIf Command1.BackColor = &HC0C0FF Then
Command1.BackColor = &H8000000F
A = A - 1
Call PortOut(888, A)
Command1.Caption = "Turn Line 1 ON"
End If
End Sub

Private Sub Command10_Click()
If Command10.BackColor = &H8000000F Then
Command10.BackColor = &HC0C0FF
A = A + 128
Call PortOut(888, A)
Command10.Caption = "Turn Line 8 Off"
ElseIf Command10.BackColor = &HC0C0FF Then
Command10.BackColor = &H8000000F
A = A - 128
Call PortOut(888, A)
Command10.Caption = "Turn Line 8 ON"
End If
End Sub

Private Sub Command11_Click()
If Timer2.Enabled = True Then
Command11.Caption = "Blink Line 1"
Timer2.Enabled = False
Call PortOut(888, 0)
Command11.BackColor = &H8000000F
ElseIf Timer2.Enabled = False Then
Command11.Caption = "Stop"
Timer2.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Command1.BackColor = &H8000000F
Command10.BackColor = &H8000000F
Command4.BackColor = &H8000000F
Command5.BackColor = &H8000000F
Command6.BackColor = &H8000000F
Command7.BackColor = &H8000000F
Command8.BackColor = &H8000000F
Command9.BackColor = &H8000000F
Command1.Caption = "Turn Line 1 ON"
Command4.Caption = "Turn Line 2 ON"
Command5.Caption = "Turn Line 3 ON"
Command6.Caption = "Turn Line 4 ON"
Command7.Caption = "Turn Line 5 ON"
Command8.Caption = "Turn Line 6 ON"
Command9.Caption = "Turn Line 7 ON"
Command10.Caption = "Turn Line 8 ON"
A = 0
Call PortOut(888, 0)
End Sub

Private Sub Command3_Click()
save_all
Unload frmSplash
Unload Me
End Sub

Private Sub Command4_Click()
If Command4.BackColor = &H8000000F Then
Command4.BackColor = &HC0C0FF
A = A + 2
Call PortOut(888, A)
Command4.Caption = "Turn Line 2 Off"
ElseIf Command4.BackColor = &HC0C0FF Then
Command4.BackColor = &H8000000F
A = A - 2
Call PortOut(888, A)
Command4.Caption = "Turn Line 2 ON"
End If
End Sub

Private Sub Command5_Click()
If Command5.BackColor = &H8000000F Then
Command5.BackColor = &HC0C0FF
A = A + 4
Call PortOut(888, A)
Command5.Caption = "Turn Line 3 Off"
ElseIf Command5.BackColor = &HC0C0FF Then
Command5.BackColor = &H8000000F
A = A - 4
Call PortOut(888, A)
Command5.Caption = "Turn Line 3 ON"
End If
End Sub

Private Sub Command6_Click()
If Command6.BackColor = &H8000000F Then
Command6.BackColor = &HC0C0FF
A = A + 8
Call PortOut(888, A)
Command6.Caption = "Turn Line 4 Off"
ElseIf Command6.BackColor = &HC0C0FF Then
Command6.BackColor = &H8000000F
A = A - 8
Call PortOut(888, A)
Command6.Caption = "Turn Line 4 ON"
End If
End Sub

Private Sub Command7_Click()
If Command7.BackColor = &H8000000F Then
Command7.BackColor = &HC0C0FF
A = A + 16
Call PortOut(888, A)
Command7.Caption = "Turn Line 5 Off"
ElseIf Command7.BackColor = &HC0C0FF Then
Command7.BackColor = &H8000000F
A = A - 16
Call PortOut(888, A)
Command7.Caption = "Turn Line 5 ON"
End If
End Sub

Private Sub Command8_Click()
If Command8.BackColor = &H8000000F Then
Command8.BackColor = &HC0C0FF
A = A + 32
Call PortOut(888, A)
Command8.Caption = "Turn Line 6 Off"
ElseIf Command8.BackColor = &HC0C0FF Then
Command8.BackColor = &H8000000F
A = A - 32
Call PortOut(888, A)
Command8.Caption = "Turn Line 6 ON"
End If
End Sub

Private Sub Command9_Click()
If Command9.BackColor = &H8000000F Then
Command9.BackColor = &HC0C0FF
A = A + 64
Call PortOut(888, A)
Command9.Caption = "Turn Line 7 Off"
ElseIf Command9.BackColor = &HC0C0FF Then
Command9.BackColor = &H8000000F
A = A - 64
Call PortOut(888, A)
Command9.Caption = "Turn Line 7 ON"
End If
End Sub

Private Sub Form_Load()
frmSplash.Show
Form1.Visible = False
Label1.Caption = Now
Text1.Text = GetSetting("IO", "Text", "1")
Text2.Text = GetSetting("IO", "Text", "2")
Text3.Text = GetSetting("IO", "Text", "3")
Text4.Text = GetSetting("IO", "Text", "4")
Text5.Text = GetSetting("IO", "Text", "5")
Text6.Text = GetSetting("IO", "Text", "6")
Text7.Text = GetSetting("IO", "Text", "7")
Text8.Text = GetSetting("IO", "Text", "8")
If Text1.Text = "" Then
Text1.Text = "Light1"
End If
Dim B As Integer
For B = 1 To 12
Combo1.AddItem B
Combo4.AddItem B
Combo18.AddItem B
Combo15.AddItem B
Combo30.AddItem B
Combo27.AddItem B
Combo42.AddItem B
Combo39.AddItem B
Combo12.AddItem B
Combo9.AddItem B
Combo24.AddItem B
Combo21.AddItem B
Combo36.AddItem B
Combo33.AddItem B
Combo48.AddItem B
Combo45.AddItem B
Next B
Loadup_Minutes Combo2
Loadup_Minutes Combo5
Loadup_Minutes Combo17
Loadup_Minutes Combo14
Loadup_Minutes Combo29
Loadup_Minutes Combo26
Loadup_Minutes Combo41
Loadup_Minutes Combo38
Loadup_Minutes Combo11
Loadup_Minutes Combo8
Loadup_Minutes Combo23
Loadup_Minutes Combo20
Loadup_Minutes Combo35
Loadup_Minutes Combo32
Loadup_Minutes Combo47
Loadup_Minutes Combo44
Combo3.AddItem "AM"
Combo3.AddItem "PM"
Combo6.AddItem "AM"
Combo6.AddItem "PM"
Combo16.AddItem "AM"
Combo16.AddItem "PM"
Combo13.AddItem "AM"
Combo13.AddItem "PM"
Combo28.AddItem "AM"
Combo28.AddItem "PM"
Combo25.AddItem "AM"
Combo25.AddItem "PM"
Combo40.AddItem "AM"
Combo40.AddItem "PM"
Combo37.AddItem "AM"
Combo37.AddItem "PM"
Combo10.AddItem "AM"
Combo10.AddItem "PM"
Combo7.AddItem "AM"
Combo7.AddItem "PM"
Combo22.AddItem "AM"
Combo22.AddItem "PM"
Combo19.AddItem "AM"
Combo19.AddItem "PM"
Combo34.AddItem "AM"
Combo34.AddItem "PM"
Combo31.AddItem "AM"
Combo31.AddItem "PM"
Combo46.AddItem "AM"
Combo46.AddItem "PM"
Combo43.AddItem "AM"
Combo43.AddItem "PM"
End Sub

Private Sub Form_Resize()
Command2.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
save_all
Unload frmSplash
Unload Me
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
save_all
Unload frmSplash
Unload Me
End Sub

Private Sub mnuHelp_Click()
Form2.Show
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Now
Label18.Caption = Str(Time)
End Sub

Public Sub save_all()
Call PortOut(888, 0)
SaveSetting "IO", "Text", "1", Text1.Text
SaveSetting "IO", "Text", "2", Text2.Text
SaveSetting "IO", "Text", "3", Text3.Text
SaveSetting "IO", "Text", "4", Text4.Text
SaveSetting "IO", "Text", "5", Text5.Text
SaveSetting "IO", "Text", "6", Text6.Text
SaveSetting "IO", "Text", "7", Text7.Text
SaveSetting "IO", "Text", "8", Text8.Text
End Sub

Private Sub Timer2_Timer()
B = B + 1
If B = 1 Then
Call PortOut(888, 1)
ElseIf B = 2 Then
Call PortOut(888, 0)
B = 0
End If
End Sub

Private Sub tmr1_Timer()
If Combo1.Text & ":" & Combo2.Text & ":" & "00" & " " & Combo3.Text = Str(Time) Then
Command1.BackColor = &HC0C0FF
A = A + 1
Call PortOut(888, A)
Command1.Caption = "Turn Line 1 Off"
End If
If Combo4.Text & ":" & Combo5.Text & ":" & "00" & " " & Combo6.Text = Str(Time) Then
Command1.BackColor = &H8000000F
A = A - 1
Call PortOut(888, A)
Command1.Caption = "Turn Line 1 ON"
End If
End Sub

Private Sub tmr2_Timer()
If Combo18.Text & ":" & Combo17.Text & ":" & "00" & " " & Combo16.Text = Str(Time) Then
Command4.BackColor = &HC0C0FF
A = A + 2
Call PortOut(888, A)
Command4.Caption = "Turn Line 2 Off"
End If
If Combo15.Text & ":" & Combo14.Text & ":" & "00" & " " & Combo13.Text = Str(Time) Then
Command4.BackColor = &H8000000F
A = A - 2
Call PortOut(888, A)
Command4.Caption = "Turn Line 2 ON"
End If
End Sub

Private Sub tmr3_Timer()
If Combo30.Text & ":" & Combo29.Text & ":" & "00" & " " & Combo28.Text = Str(Time) Then
Command5.BackColor = &HC0C0FF
A = A + 4
Call PortOut(888, A)
Command5.Caption = "Turn Line 3 Off"
End If
If Combo27.Text & ":" & Combo26.Text & ":" & "00" & " " & Combo25.Text = Str(Time) Then
Command5.BackColor = &H8000000F
A = A - 4
Call PortOut(888, A)
Command5.Caption = "Turn Line 3 ON"
End If
End Sub

Private Sub tmr4_Timer()
If Combo42.Text & ":" & Combo41.Text & ":" & "00" & " " & Combo40.Text = Str(Time) Then
Command6.BackColor = &HC0C0FF
A = A + 8
Call PortOut(888, A)
Command6.Caption = "Turn Line 4 Off"
End If
If Combo39.Text & ":" & Combo38.Text & ":" & "00" & " " & Combo37.Text = Str(Time) Then
Command6.BackColor = &H8000000F
A = A - 8
Call PortOut(888, A)
Command6.Caption = "Turn Line 4 ON"
End If
End Sub

Private Sub tmr5_Timer()
If Combo12.Text & ":" & Combo11.Text & ":" & "00" & " " & Combo10.Text = Str(Time) Then
Command7.BackColor = &HC0C0FF
A = A + 16
Call PortOut(888, A)
Command7.Caption = "Turn Line 5 Off"
End If
If Combo9.Text & ":" & Combo8.Text & ":" & "00" & " " & Combo7.Text = Str(Time) Then
Command7.BackColor = &H8000000F
A = A - 16
Call PortOut(888, A)
Command7.Caption = "Turn Line 5 ON"
End If
End Sub

Private Sub tmr6_Timer()
If Combo24.Text & ":" & Combo23.Text & ":" & "00" & " " & Combo22.Text = Str(Time) Then
Command8.BackColor = &HC0C0FF
A = A + 32
Call PortOut(888, A)
Command8.Caption = "Turn Line 6 Off"
End If
If Combo21.Text & ":" & Combo20.Text & ":" & "00" & " " & Combo19.Text = Str(Time) Then
Command8.BackColor = &H8000000F
A = A - 32
Call PortOut(888, A)
Command8.Caption = "Turn Line 6 ON"
End If
End Sub

Private Sub tmr7_Timer()
If Combo36.Text & ":" & Combo35.Text & ":" & "00" & " " & Combo34.Text = Str(Time) Then
Command9.BackColor = &HC0C0FF
A = A + 64
Call PortOut(888, A)
Command9.Caption = "Turn Line 7 Off"
End If
If Combo33.Text & ":" & Combo32.Text & ":" & "00" & " " & Combo31.Text = Str(Time) Then
Command9.BackColor = &H8000000F
A = A - 64
Call PortOut(888, A)
Command9.Caption = "Turn Line 7 ON"
End If
End Sub

Private Sub tmr8_Timer()
If Combo48.Text & ":" & Combo47.Text & ":" & "00" & " " & Combo46.Text = Str(Time) Then
Command10.BackColor = &HC0C0FF
A = A + 128
Call PortOut(888, A)
Command10.Caption = "Turn Line 8 Off"
End If
If Combo45.Text & ":" & Combo44.Text & ":" & "00" & " " & Combo43.Text = Str(Time) Then
Command10.BackColor = &H8000000F
A = A - 128
Call PortOut(888, A)
Command10.Caption = "Turn Line 8 ON"
End If
End Sub
