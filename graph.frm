VERSION 5.00
Begin VB.Form Graph2D 
   Caption         =   "Graph2D, Made by S.M.Asaduzzaman, M.S., Dept. Of Math., University Of Dhaka."
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame trigfram 
      Height          =   975
      Left            =   480
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox n2 
         Height          =   285
         Index           =   3
         Left            =   7320
         TabIndex        =   104
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox a1 
         Height          =   285
         Index           =   5
         Left            =   6600
         TabIndex        =   103
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox n2 
         Height          =   285
         Index           =   1
         Left            =   6000
         TabIndex        =   102
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox a1 
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   101
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox n2 
         Height          =   285
         Index           =   0
         Left            =   4680
         TabIndex        =   100
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox a1 
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   99
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox n1 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   98
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox a1 
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   97
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox n1 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   96
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox a1 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   95
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox n1 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   94
         Text            =   "0"
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox a1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   93
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label x1 
         Caption         =   "x "
         Height          =   375
         Index           =   5
         Left            =   7680
         TabIndex        =   117
         Top             =   360
         Width           =   375
      End
      Begin VB.Label sinnx 
         Caption         =   "Cot"
         Height          =   255
         Index           =   5
         Left            =   6960
         TabIndex        =   116
         Top             =   360
         Width           =   375
      End
      Begin VB.Label x1 
         Caption         =   "x +"
         Height          =   375
         Index           =   4
         Left            =   6360
         TabIndex        =   115
         Top             =   360
         Width           =   375
      End
      Begin VB.Label sinnx 
         Caption         =   "Sec"
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   114
         Top             =   360
         Width           =   375
      End
      Begin VB.Label x1 
         Caption         =   "x +"
         Height          =   375
         Index           =   3
         Left            =   5040
         TabIndex        =   113
         Top             =   360
         Width           =   375
      End
      Begin VB.Label sinnx 
         Caption         =   "CSC"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   112
         Top             =   360
         Width           =   375
      End
      Begin VB.Label x1 
         Caption         =   "x +"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   111
         Top             =   360
         Width           =   375
      End
      Begin VB.Label sinnx 
         Caption         =   "tan"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   110
         Top             =   360
         Width           =   375
      End
      Begin VB.Label x1 
         Caption         =   "x +"
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   109
         Top             =   360
         Width           =   375
      End
      Begin VB.Label sinnx 
         Caption         =   "Cos"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   108
         Top             =   360
         Width           =   375
      End
      Begin VB.Label a1p 
         Caption         =   "+"
         Height          =   255
         Left            =   1200
         TabIndex        =   107
         Top             =   360
         Width           =   255
      End
      Begin VB.Label x1 
         Caption         =   "x"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   106
         Top             =   360
         Width           =   375
      End
      Begin VB.Label sinnx 
         Caption         =   "Sin"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   105
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame sixteencol 
      Caption         =   "SixteenColor"
      Height          =   3735
      Left            =   120
      TabIndex        =   73
      Top             =   3840
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox selectedcol 
         Height          =   495
         Left            =   2160
         TabIndex        =   91
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox col16 
         Height          =   495
         Left            =   3000
         TabIndex        =   89
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox col15 
         BackColor       =   &H0080FFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   88
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox col14 
         BackColor       =   &H00FF00FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   87
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox col13 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   120
         TabIndex        =   86
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox col12 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Left            =   3000
         TabIndex        =   85
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox col11 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   2040
         TabIndex        =   84
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox col10 
         BackColor       =   &H00FF8080&
         Height          =   495
         Left            =   1080
         TabIndex        =   83
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox col9 
         BackColor       =   &H8000000C&
         Height          =   495
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox col8 
         BackColor       =   &H80000016&
         Height          =   495
         Left            =   3000
         TabIndex        =   81
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox col7 
         BackColor       =   &H00C0FFC0&
         Height          =   495
         Left            =   2040
         TabIndex        =   80
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox col6 
         BackColor       =   &H00C0C0FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   79
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox col5 
         BackColor       =   &H00404080&
         Height          =   495
         Left            =   120
         TabIndex        =   78
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox col4 
         BackColor       =   &H00808000&
         Height          =   495
         Left            =   3000
         TabIndex        =   77
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox col3 
         BackColor       =   &H00008000&
         Height          =   495
         Left            =   2040
         TabIndex        =   76
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox col2 
         BackColor       =   &H80000002&
         Height          =   495
         Left            =   1080
         TabIndex        =   75
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox col1 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SELECTED COLOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   120
         TabIndex        =   90
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame RGBCol 
      Caption         =   "RGBCOLOR"
      Height          =   1695
      Left            =   120
      TabIndex        =   66
      Top             =   3840
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox b 
         Height          =   375
         Left            =   1800
         TabIndex        =   72
         Text            =   "0"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox g 
         Height          =   375
         Left            =   1800
         TabIndex        =   71
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox r 
         Height          =   375
         Left            =   1800
         TabIndex        =   68
         Text            =   "0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label BLUE 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "BLUE"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label GREEN 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "GREEN"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label RED 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "RED"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame mousept 
      Caption         =   "Co-ordinate at Mouse"
      Height          =   735
      Left            =   120
      TabIndex        =   61
      Top             =   3000
      Width           =   3975
      Begin VB.Label mouseptyval 
         Height          =   375
         Left            =   2520
         TabIndex        =   65
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label mousepty 
         Caption         =   "Y="
         Height          =   375
         Left            =   2040
         TabIndex        =   64
         Top             =   240
         Width           =   375
      End
      Begin VB.Label mouseptxval 
         Height          =   375
         Left            =   480
         TabIndex        =   63
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label mouseptx 
         Caption         =   "X="
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame rationalfram 
      Height          =   855
      Left            =   480
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   9735
      Begin VB.TextBox de6 
         Height          =   285
         Left            =   5160
         TabIndex        =   59
         Text            =   "0"
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox de5 
         Height          =   285
         Left            =   4080
         TabIndex        =   57
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox de4 
         Height          =   285
         Left            =   3000
         TabIndex        =   55
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox de3 
         Height          =   285
         Left            =   1920
         TabIndex        =   53
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox de2 
         Height          =   285
         Left            =   960
         TabIndex        =   51
         Text            =   "0"
         Top             =   450
         Width           =   615
      End
      Begin VB.TextBox nu1 
         Height          =   285
         Left            =   0
         TabIndex        =   50
         Text            =   "0"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox de1 
         Height          =   285
         Left            =   0
         TabIndex        =   48
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox nu6 
         Height          =   285
         Left            =   5160
         TabIndex        =   46
         Text            =   "0"
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox nu5 
         Height          =   285
         Left            =   4080
         TabIndex        =   44
         Text            =   "0"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox fnu4 
         Height          =   285
         Left            =   3000
         TabIndex        =   42
         Text            =   "0"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox fnu3 
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Text            =   "0"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox fnu2 
         Height          =   285
         Left            =   960
         TabIndex        =   38
         Text            =   "0"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label l6 
         Caption         =   "x^5"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   60
         Top             =   480
         Width           =   495
      End
      Begin VB.Label l5 
         Caption         =   "x^4+"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   58
         Top             =   480
         Width           =   375
      End
      Begin VB.Label l4 
         Caption         =   "x^3+"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   56
         Top             =   480
         Width           =   375
      End
      Begin VB.Label l3 
         Caption         =   "x^2+"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   54
         Top             =   480
         Width           =   495
      End
      Begin VB.Label l2 
         Caption         =   "x+"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   52
         Top             =   480
         Width           =   375
      End
      Begin VB.Label l1 
         Caption         =   "+"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   49
         Top             =   480
         Width           =   255
      End
      Begin VB.Label l6 
         Caption         =   "x^5"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   47
         Top             =   0
         Width           =   495
      End
      Begin VB.Label l5 
         Caption         =   "x^4+"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   45
         Top             =   0
         Width           =   375
      End
      Begin VB.Label l4 
         Caption         =   "x^3+"
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   43
         Top             =   0
         Width           =   375
      End
      Begin VB.Label l3 
         Caption         =   "x^2+"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   41
         Top             =   0
         Width           =   495
      End
      Begin VB.Label l2 
         Caption         =   "x+"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   39
         Top             =   0
         Width           =   375
      End
      Begin VB.Label l1 
         Caption         =   "+"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   37
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9600
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.TextBox fx10 
      Height          =   285
      Left            =   9000
      TabIndex        =   34
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx9 
      Height          =   285
      Left            =   8160
      TabIndex        =   32
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx8 
      Height          =   285
      Left            =   7200
      TabIndex        =   30
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx7 
      Height          =   285
      Left            =   6360
      TabIndex        =   28
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx6 
      Height          =   285
      Left            =   5400
      TabIndex        =   26
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx5 
      Height          =   285
      Left            =   4440
      TabIndex        =   24
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx4 
      Height          =   285
      Left            =   3600
      TabIndex        =   23
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
      Height          =   855
      Left            =   2880
      TabIndex        =   21
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Domain 
      Caption         =   "Domain"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   2535
      Begin VB.TextBox ydom 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Text            =   "10"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox xdom 
         Height          =   285
         Left            =   600
         TabIndex        =   19
         Text            =   "-10"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Max="
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Min="
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Scale 
      Caption         =   "Scale"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   2535
      Begin VB.TextBox yscale 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Text            =   "1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox xscale 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Text            =   "1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "YScale="
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "XScale="
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Plot"
      Height          =   855
      Left            =   2880
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox fx3 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx2 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx1 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox fx0 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox outbox 
      BackColor       =   &H00FFFFFF&
      Height          =   6645
      Left            =   4200
      MousePointer    =   2  'Cross
      ScaleHeight     =   6585
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   1080
      Width           =   7000
   End
   Begin VB.Label Label9 
      Caption         =   "x^10"
      Height          =   255
      Index           =   6
      Left            =   9600
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "x^9+"
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "x^8+"
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "x^7+"
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "x^6+"
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "x^5+"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "x^4+"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "x^3+"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "x^2+"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "x+"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label fx 
      Caption         =   "F(x)="
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu Types_Of_Graph 
      Caption         =   "&Types_Of_Graph"
      Begin VB.Menu Polynomial_function 
         Caption         =   "&Polynomial function"
      End
      Begin VB.Menu Rational_Function 
         Caption         =   "&Rational _Function"
      End
      Begin VB.Menu Trigonometric_function 
         Caption         =   "&Trigonometric_function"
      End
      Begin VB.Menu Trigonometric_Rational_function 
         Caption         =   "Trigonometric_Rational_function"
      End
   End
   Begin VB.Menu Color 
      Caption         =   "&Color"
      Begin VB.Menu RGBColor 
         Caption         =   "&RGBColor"
      End
      Begin VB.Menu SixteenColor 
         Caption         =   "&SixteenColor"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About_Author 
         Caption         =   "&About_Author"
      End
      Begin VB.Menu About_Software 
         Caption         =   "About_&Software"
      End
   End
End
Attribute VB_Name = "Graph2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Double
Dim Y As Double
Dim xscal As Double
Dim yscal As Double
Dim xdomain As Double
Dim ydomain As Double
Dim xsetcoordinate As Double
Dim ysetcoordinate As Double
Dim typeofgraph As Double
Dim mousepx As Double
Dim mousepy As Double
Dim rgbred As Integer
Dim rgbgreen As Integer
Dim rgbblue As Integer
Dim colorval As Integer
Dim sixteencols As Integer

Private Sub About_Author_Click()
Dim msg
msg = MsgBox( _
"Name: S.M.ASADUZZAMAN,M.S.,DEPT. OF MATH.,ROOM NO. 1027, SHAHIDULLAH HALL, UNIVERSITY OF DHAKA", _
vbOKOnly, "AUTHOR INFORMATION")
End Sub

Private Sub About_Software_Click()
Dim msg1
msg1 = MsgBox("THIS SOFTWARE MAKE GRAPH OF MATHEMATICAL FUNCTION", _
vbOKOnly, "SOFTWARE INFORMATION")
End Sub

Private Sub Clear_Click()
' Clear background of the out put screen
outbox.BackColor = RGB(255, 255, 255)
End Sub



Private Sub col1_click()
sixteencols = 0
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col10_click()
sixteencols = 9
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col11_click()
sixteencols = 10
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col12_click()
sixteencols = 11
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col13_click()
sixteencols = 12
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col14_click()
sixteencols = 13
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col15_click()
sixteencols = 14
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col16_click()
sixteencols = 15
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col2_click()
sixteencols = 1
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col3_click()
sixteencols = 2
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col4_click()
sixteencols = 3
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col5_click()
sixteencols = 4
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col6_click()
sixteencols = 5
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col7_click()
sixteencols = 6
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col8_click()
sixteencols = 7
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub col9_click()
sixteencols = 8
selectedcol.BackColor = QBColor(sixteencols)
End Sub

Private Sub Command1_Click()
' scale calculation
xscal = Val(xscale.Text)
yscal = Val(yscale.Text)
'dom calculation
xdomain = Val(xdom.Text)
ydomain = Val(ydom.Text)
'Drawing dept of line
outbox.DrawWidth = 1
' Drawing x-axis
outbox.Line (0, outbox.ScaleHeight / 2)-(outbox.ScaleWidth _
, outbox.ScaleHeight / 2), RGB(0, 0, 0)
'Drawing Y-axis
outbox.Line (outbox.ScaleWidth / 2, 0)-(outbox.ScaleWidth / 2 _
, outbox.ScaleHeight), RGB(0, 0, 0)
' Drawing graph
outbox.DrawWidth = 1
'set co-ordinate
xsetcoordinate = outbox.ScaleWidth * xscal / (ydomain - xdomain)
ysetcoordinate = outbox.ScaleWidth * yscal / (ydomain - xdomain)
'For x = -outbox.ScaleWidth / (2 * xscal) To outbox.ScaleWidth / _
(2 * yscal) Step 0.1
For X = xdomain To ydomain Step 0.0051
' codes for rgb color
If colorval = 1 Then
' assign rgb colors
rgbred = Val(r.Text)
rgbgreen = Val(g.Text)
rgbblue = Val(b.Text)
' code for polynomial function and rgb color
If typeofgraph = 1 Then
Y = Val(fx0 + fx1 * X + fx2 * X ^ 2 + fx3 * X ^ 3 + fx4 * X ^ 4 _
+ fx5 * X ^ 5 + fx6 * X ^ 6 + fx7 * X ^ 7 + fx8 * X ^ 8 + _
fx9 * X ^ 9 + fx10 * X ^ 10) / 10
' code for rational function and rgb color
ElseIf typeofgraph = 2 Then
Y = Val(nu1 + nu2 * X + nu3 * X ^ 2 + nu4 * X ^ 3 + nu5 * X ^ 4 + nu6 * X ^ 5) _
/ Val(de1 + de2 * X + de3 * X ^ 2 + de4 * X ^ 3 + de5 * X ^ 4 + de6 * X ^ 6)
' codes for trigonometric functionsa and rgbcolor
ElseIf typeofgraph = 3 Then
Y = Val(a1(0) * Sin(n1(0) * X) + a1(1) * Cos(n1(1) * X) _
+ a1(2) * Tan(n1(2) * X) + a1(3) / Sin(n2(0) * X) _
+ a1(4) / Cos(n2(1) * X) + a1(5) / Tan(n2(3) * X))
End If
outbox.PSet (outbox.ScaleWidth / 2 + X * xsetcoordinate, _
outbox.ScaleHeight / 2 - Y * ysetcoordinate), RGB(rgbred, rgbgreen, rgbblue)
' codes for sixteen colors
ElseIf colorval = 2 Then
' code for polynomial function and sixteen  color
If typeofgraph = 1 Then
Y = Val(fx0 + fx1 * X + fx2 * X ^ 2 + fx3 * X ^ 3 + fx4 * X ^ 4 _
+ fx5 * X ^ 5 + fx6 * X ^ 6 + fx7 * X ^ 7 + fx8 * X ^ 8 + _
fx9 * X ^ 9 + fx10 * X ^ 10) / 10
' code for rational function and sixteen color
ElseIf typeofgraph = 2 Then
Y = Val(nu1 + nu2 * X + nu3 * X ^ 2 + nu4 * X ^ 3 + nu5 * X ^ 4 + nu6 * X ^ 5) _
/ Val(de1 + de2 * X + de3 * X ^ 2 + de4 * X ^ 3 + de5 * X ^ 4 + de6 * X ^ 6)
' codes for trigonometric functionsa and sixteencolor
ElseIf typeofgraph = 3 Then
Y = Val(a1(0) * Sin(n1(0) * X) + a1(1) * Cos(n1(1) * X) _
+ a1(2) * Tan(n1(2) * X) + a1(3) / Sin(n2(0) * X) _
+ a1(4) / Cos(n2(1) * X) + a1(5) / Tan(n2(3) * X))
End If
outbox.PSet (outbox.ScaleWidth / 2 + X * xsetcoordinate, _
outbox.ScaleHeight / 2 - Y * ysetcoordinate), QBColor(sixteencols)
' codes for black color
Else
' code for polynomial function and black  color
If typeofgraph = 1 Then
Y = Val(fx0 + fx1 * X + fx2 * X ^ 2 + fx3 * X ^ 3 + fx4 * X ^ 4 _
+ fx5 * X ^ 5 + fx6 * X ^ 6 + fx7 * X ^ 7 + fx8 * X ^ 8 + _
fx9 * X ^ 9 + fx10 * X ^ 10) / 10
' code for rational function and sixteen color
ElseIf typeofgraph = 2 Then
Y = Val(nu1 + nu2 * X + nu3 * X ^ 2 + nu4 * X ^ 3 + nu5 * X ^ 4 + nu6 * X ^ 5) _
/ Val(de1 + de2 * X + de3 * X ^ 2 + de4 * X ^ 3 + de5 * X ^ 4 + de6 * X ^ 6)
' codes for trigonometric functionsa and black color
ElseIf typeofgraph = 3 Then
Y = Val(a1(0) * Sin(n1(0) * X) + a1(1) * Cos(n1(1) * X) _
+ a1(2) * Tan(n1(2) * X) + a1(3) / Sin(n2(0) * X) _
+ a1(4) / Cos(n2(1) * X) + a1(5) / Tan(n2(3) * X))
End If
outbox.PSet (outbox.ScaleWidth / 2 + X * xsetcoordinate, _
outbox.ScaleHeight / 2 - Y * ysetcoordinate), RGB(0, 0, 0)
End If
Next
' showing rgb color box invisiable
RGBCol.Visible = False
' showing qbcolor box invisiable
sixteencol.Visible = False
End Sub

Private Sub outbox_Click()
mousepx = outbox.CurrentX
mousepy = outbox.CurrentY
mouseptxval.Caption = mousepx
mouseptyval.Caption = mousepy
End Sub

Private Sub Polynomial_function_Click()
'type of graph 1 for polynomial
typeofgraph = 1
' desiable rational function
rationalfram.Visible = False
' showing labels
fx.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label9(0).Visible = True
Label9(1).Visible = True
Label9(2).Visible = True
Label9(3).Visible = True
Label9(4).Visible = True
Label9(5).Visible = True
Label9(6).Visible = True

'showing text box
fx0.Visible = True
fx1.Visible = True
fx2.Visible = True
fx3.Visible = True
fx4.Visible = True
fx5.Visible = True
fx6.Visible = True
fx7.Visible = True
fx8.Visible = True
fx9.Visible = True
fx10.Visible = True

'text box initialization
fx0.Text = 0
fx1.Text = 0
fx2.Text = 0
fx3.Text = 0
fx4.Text = 0
fx5.Text = 0
fx6.Text = 0
fx7.Text = 0
fx8.Text = 0
fx9.Text = 0
fx10.Text = 0
End Sub

Private Sub Rational_Function_Click()
' trig fram invisiable
trigfram.Visible = False
'type of graph 2 for rational
typeofgraph = 2
rationalfram.Visible = True
End Sub

Private Sub RGBColor_Click()
' for rgb colorval is 1
colorval = 1
' sixteen box invisiable
sixteencol.Visible = False
' show rgb color
RGBCol.Visible = True
' assign rgb colors
rgbred = Val(r.Text)
rgbgreen = Val(g.Text)
rgbblue = Val(b.Text)
End Sub

Private Sub SixteenColor_Click()
' for sixteen colorval is 2
colorval = 2
' show rgb color invisiable
RGBCol.Visible = False
' sixteen box visiable
sixteencol.Visible = True
'colors assign
col1.BackColor = QBColor(0)
col2.BackColor = QBColor(1)
col3.BackColor = QBColor(2)
col4.BackColor = QBColor(3)
col5.BackColor = QBColor(4)
col6.BackColor = QBColor(5)
col7.BackColor = QBColor(6)
col8.BackColor = QBColor(7)
col9.BackColor = QBColor(8)
col10.BackColor = QBColor(9)
col11.BackColor = QBColor(10)
col12.BackColor = QBColor(11)
col13.BackColor = QBColor(12)
col14.BackColor = QBColor(13)
col15.BackColor = QBColor(14)
col16.BackColor = QBColor(15)
End Sub


Private Sub Trigonometric_function_Click()
'invisiable rationnalframe
rationalfram.Visible = False
'type of graph 3 for rational
typeofgraph = 3
trigfram.Visible = True
fx.Visible = True
End Sub

