VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "LMP Calculator V4.0"
   ClientHeight    =   7050
   ClientLeft      =   5940
   ClientTop       =   2430
   ClientWidth     =   4410
   ControlBox      =   0   'False
   ForeColor       =   &H00FF8080&
   Icon            =   "cal41.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "cal41.frx":0442
   ScaleHeight     =   7050
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   5760
      TabIndex        =   50
      Text            =   "Text5"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5880
      TabIndex        =   49
      Text            =   "Text4"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4800
      TabIndex        =   48
      Text            =   "Text3"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   47
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   3360
      TabIndex        =   46
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   3360
      TabIndex        =   45
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "f(A)"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   44
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   43
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   42
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   41
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "10^"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   2640
      TabIndex        =   40
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   1920
      TabIndex        =   39
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   1920
      TabIndex        =   38
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1920
      TabIndex        =   37
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "lg(x)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1920
      TabIndex        =   36
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1200
      TabIndex        =   35
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Xn"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   34
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÖC"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   480
      TabIndex        =   33
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   32
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X²"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   480
      TabIndex        =   31
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "cmpx"
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   1920
      TabIndex        =   29
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "vec"
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "rad"
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "deg"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "M+"
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "shift"
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   22
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   21
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "shift"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2640
      TabIndex        =   19
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3360
      TabIndex        =   18
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   3360
      TabIndex        =   16
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2640
      TabIndex        =   15
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3360
      TabIndex        =   14
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton operators 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2640
      TabIndex        =   12
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   2640
      TabIndex        =   11
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   1920
      TabIndex        =   10
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1200
      TabIndex        =   9
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   480
      TabIndex        =   8
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   1920
      TabIndex        =   7
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1920
      TabIndex        =   4
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Numpanel 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim m As Double
Dim comwork As String
Dim a1 As Double
Dim a2 As Double
Dim b2 As Double
Dim b1 As Double
Dim c1 As Double
Dim c2 As Double
Dim func As Integer
Const pi = 3.141592654
Dim mode As Integer
Dim f As Double
Dim fx As Double
Dim work As String
Dim num As Boolean
Dim shift As Boolean
Dim alpha As Boolean


Private Sub Command1_Click(Index As Integer)
Text1(5) = ""
Select Case Index
Case 1
End
Case 3
If shift = True Then
shift = False
Text1(0).Visible = False
Else
shift = True
Text1(0).Visible = True
End If
Case 0
If shift = False Then txt = Val(txt) * -1 Else txt = pi
Case 2
work = ""
txt = ""
Text1(5) = ""
a = ""
 b = ""
 c = ""
 d = ""
 comwork = ""
 a1 = 0
 a2 = 0
 b1 = 0
 func = 0
 b2 = 0
 c1 = 0
 c2 = 0
Case 1
Case 5
If shift = False Then
txt = m
Else
c = fx / f
b = f & " numbers were entered.Total is " & fx & ".Average is " & c
d = MsgBox(b & ".", vbOKOnly + vbInformation, "Data summary")
fx = 0
f = 0
End If

Case 4
If shift = True Then
If alpha = True Then
alpha = False
Text1(2).Visible = True
Text1(3).Visible = False
Else
If alpha = False Then
alpha = True
Text1(2).Visible = False
Text1(3).Visible = True
End If
End If
End If
If shift = False Then
If mode = 0 Then
mode = 1
Text1(6).Visible = True
Else
If mode = 1 Then
mode = 0
Text1(6).Visible = False
End If
End If
End If

Case 6
txt = Val(txt) * 100

Case 7
If shift = False Then txt = Val(txt) * Val(txt) Else txt = Val(txt) * Val(txt) * Val(txt)

Case 8
If shift = False Then txt = 1 / Val(txt) Else txt = Rnd(1)

Case 9
If shift = False Then
If Val(txt) >= 0 Then txt = Sqr(Val(txt))
If Val(txt) < 0 Then txt = Sqr(-Val(txt)) & "i"
Else
txt = Val(txt) ^ (1 / 3)
End If

Case 11
If shift = True Then
If Val(txt) = 1 Then txt = 1.618
If Val(txt) = 2 Then txt = 8.314472
If Val(txt) = 3 Then txt = 9.80665
If Val(txt) = 4 Then txt = 8.854187817 * (10 ^ -12)
End If
If shift = False Then
Dim g As Double
Dim base As Integer
Dim s As Double
Dim result As String
Dim e As String

g = Val(txt)
txt = ""
base = InputBox("Enter the base of the digit system. Bases greater than 10 cannot be displayed correctly. ", "Enter the base", "2")
If base > 10 Then
e = MsgBox("digits of bases greater than 10 couldn't be displayed.", vbOKOnly + vbCritical, "error")
GoTo start
End If
start:

s = g
g = Fix(g / base)
e = g * base
e = s - e
result = e & result
GoTo counter

counter:
If g < base Then
txt = g & result
Else
GoTo start

End If
End If

Case 12
If shift = True Then GoTo taninverse
If shift = False Then
txt = fact(Val(txt))
Exit Sub
End If

taninverse:
b = Atn(Val(txt))
If alpha = False Then txt = Round(180 * b / pi, 4) Else txt = b

Case 13
If shift = False Then txt = Log(Val(txt)) / Log(10) Else txt = Log(Val(txt)) / Log(Exp(1))

Case 14
If shift = False Then
If alpha = False Then b = Val(txt) * pi / 180 Else b = Val(txt)
b = Sin(b)
txt = Round(b, 6)
Else
If alpha = False Then b = Val(txt) * pi / 180 Else b = Val(txt)
txt = sinh(Val(b))
End If

Case 15
If shift = False Then
If alpha = False Then b = Val(txt) * pi / 180 Else b = Val(txt)
b = Cos(b)
txt = Round(b, 6)
Else
If alpha = False Then b = Val(txt) * pi / 180 Else b = Val(txt)
txt = cosh(Val(b))
End If

Case 16
If shift = False Then
If alpha = False Then b = Val(txt) * pi / 180 Else b = Val(txt)
b = Tan(b)
txt = Round(b, 6)
Else
If alpha = False Then b = Val(txt) * pi / 180 Else b = Val(txt)
txt = tanh(Val(b))
End If

Case 17
If shift = False Then txt = 10 ^ Val(txt) Else txt = Exp(Val(txt))



End Select
End Sub




Private Sub Command2_Click()
If shift = False Then
Select Case mode
Case 1
If func = 0 Then a2 = Val(txt)
If func = 0 Then func = 1
If func = 2 Then
func = 3
txt = "arg(" & a1 & "+" & a2 & "i)"
comwork = "arg"
work = "arg"
Else
If func = 1 Then
func = 2
txt = "CmpxAns"
comwork = "z"
work = "z"
Else
If func = 3 Then
func = 1
txt = "|" & a1 & "+" & a2 & "i|"
comwork = "|z|"
work = "|z|"
End If
End If
End If
Text2 = comwork
Text3 = a1
Text4 = a2
 End Select
End If
End Sub




Private Sub Form_Load()
txt = "<=(//{ LM-348K }\\)=>"
End Sub

Private Sub Numpanel_Click(Index As Integer)
If num = False Then
txt = ""
num = True
End If
txt = txt & Numpanel(Index).Caption
End Sub

Private Sub operators_Click(Index As Integer)
num = False

Select Case work
Case ""

Case "add"
 If mode = 0 Then txt = Val(txt) + Val(a)
Case "minus"
 If mode = 0 Then txt = Val(a) - Val(txt)
Case "multiply"
 If mode = 0 Then txt = Val(a) * Val(txt)
Case "divide"
 If mode = 0 Then If a = 0 Then txt = "Are you nuts?" Else txt = Val(a) / Val(txt)
Case "power"
txt = Val(b) ^ Val(txt)
Case "halfpower"
txt = 1 / Val(txt)
txt = Val(b) ^ Val(txt)
Case "log"
txt = Log(Val(b)) / Log(Val(txt))


Case "complex"
If b1 = 0 Then
a2 = Val(txt)
If Index = 1 Then comwork = "add"
If Index = 2 Then comwork = "minus"
If Index = 4 Then comwork = "multiply"
If Index = 6 Then comwork = "divide"
End If
End Select

work = ""
a = Val(txt)

Select Case Index
Case 8
b = Val(txt)

Case 0
b = a
If shift = False Then work = "power" Else work = "halfpower"

Case 3
Text1(5) = "="
If mode = 1 Then GoTo comcount
If mode = 0 Then
If shift = False Then
txt = a
Else
txt = Val(a * 100)
End If
End If
GoTo comend

comcount:
If mode <> 1 Then Exit Sub
b2 = Val(txt)
Text2 = a1
Text3 = a2
Text4 = b1
Text5 = b2
If comwork = "z" Then GoTo zcount
If comwork = "add" Then
c1 = (a1 + b1)
c2 = (a2 + b2)
End If
If comwork = "minus" Then
c1 = (a1 - b1)
c2 = (a2 - b2)
End If
If comwork = "multiply" Then
c1 = (a1 * b1 - a2 * b2)
c2 = (a1 * b2 + b1 * a2)
End If
If comwork = "divide" Then
c1 = Round((a1 * b1 + a2 * b2) / (b1 * b1 + b2 * b2), 6)
c2 = Round((a2 * b1 - b2 * a1) / (b1 * b1 + b2 * b2), 6)
End If
If c2 > 0 Then txt = c1 & "+" & c2 & "i"
If c2 < 0 Then txt = c1 & c2 & "i"
If c2 = 0 Then txt = c1
If comwork = "arg" Then
b = Atn(a2 / a1)
If alpha = False Then txt = Round(180 * b / pi, 4) Else txt = b
End If
If comwork = "|z|" Then txt = Sqr(a1 * a1 + a2 * a2)
GoTo comend

comend:
a = ""
 b = ""
 c = ""
 d = ""
 work = ""
 comwork = ""
 b1 = 0
 b2 = 0
  a1 = 0
 a2 = 0
 func = 0
Case 1
Text1(5) = "+"
If mode <> 1 Then
work = "add"
End If
Exit Sub

zcount:
a1 = c1
a2 = c2

Case 2
work = "minus"
Text1(5) = "-"
Case 4
work = "multiply"
Text1(5) = "x"
Case 6
work = "divide"
Text1(5) = "÷"

Case 7
If shift = True Then
If mode = 1 Then GoTo bypass
b = Val(txt)
work = "log"
End If
GoTo bypass
bypass:
If mode = 1 Then
Text1(5) = "i"
If a1 = 0 Then
a1 = Val(txt)
work = "complex"
End If
If a1 <> 0 And a2 <> 0 Then
b1 = Val(txt)
work = "complex"
End If
End If
Text2 = a1
Text3 = a2
Text4 = b1
Text5 = b2
Case 5
If shift = False Then
m = Val(txt)
Text1(1).Visible = True
Else
fx = Val(txt) + fx
f = f + 1
End If
End Select

End Sub

Private Sub txt_DblClick()
MsgBox "LMP Calculator V4.x is designed by Lasika Malshan Peiris- (mail to: lmpeiris@gmail.com, mobile: +94 077 5525110, web site: http://lmpeiris.hi5.com)." & Chr$(13) & "Version " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation + vbOKOnly, "LMP Calculator V4.x"
End Sub
