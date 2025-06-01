VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "LMP CALCULATOR V3.1"
   ClientHeight    =   3705
   ClientLeft      =   225
   ClientTop       =   3330
   ClientWidth     =   6660
   ControlBox      =   0   'False
   FillColor       =   &H80000016&
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0442
   MousePointer    =   10  'Up Arrow
   ScaleHeight     =   3705
   ScaleWidth      =   6660
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "x^½"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   3840
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FF80FF&
      Caption         =   "Xn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   16
      Left            =   5280
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FF80FF&
      Caption         =   "x!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   4560
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   4560
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   4560
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "con"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   6000
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "lg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3840
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FF80FF&
      Caption         =   "set"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF80FF&
      Caption         =   "shift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF80FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2400
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF80FF&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "sqr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3840
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF80FF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF80FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF80FF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF80FF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2400
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2400
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1680
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   960
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   240
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1680
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   960
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1680
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   960
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label lb1 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   60
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fx As Double
Dim f As Double
Dim memo As String
Dim a As Double
Dim c As Double
Dim M As String
Dim N As String

Dim ff As String
Const z = 22 / 7
Dim shift As Integer
Dim B As String

Private Sub Command1_Click(Index As Integer)
txt1.Text = txt1.Text & Index
End Sub



Private Sub Command10_Click()

a = Val(txt1)
B = 1
GoTo counter
counter:
B = a * (a - 1) * B
a = a - 2
GoTo processer
processer:
If a < 2 Then
txt1 = B
Else
GoTo counter
End If
End Sub

Private Sub Command11_Click()
Form2.Show
End Sub

Private Sub Command12_Click()

Dim g As Double
Dim base As Double
Dim d As Double
Dim result As String
Dim e As String

g = Val(txt1)
txt1 = ""
base = InputBox("Enter the base of the digit system. Bases greater than 10 cannot be displayed correctly. ", "Enter the base")
If base > 10 Then
e = MsgBox(N & ",digits of bases greater than 10 couldn't be displayed.", vbOKOnly + vbCritical, "error")
GoTo start
End If
start:
d = g
g = Fix(g / base)
e = g * base
e = d - e
result = e & result
GoTo counter

counter:
If g < base Then
txt1 = g & result
Else
GoTo start

End If
End Sub

Private Sub Command2_Click(Index As Integer)
a = Val(txt1.Text)
txt1.Text = ""
On Error GoTo handler
handler:
Resume Next
Select Case (Index)
Case 0
B = "add"
Case 1
B = "Minus"
Case 2
B = "Multi"
Case 3
B = "Div"
Case "4"
If shift = 0 Then
txt1 = Sqr(a)
Else
a = Sqr(a)
txt1 = Sqr(a)

End If
Case "5"
If shift = 0 Then
txt1 = Log(a) / Log(10)
Else
c = InputBox("Enter log base.Decimal 10 base is the default.", "which base?")
txt1 = Log(a) / Log(c)

End If
Case "6"
If shift = 0 Then
txt1 = M
Else
c = fx / f
B = f & " numbers were entered.Total is " & fx & ".Average is " & c
ff = MsgBox(B & ".", vbOKOnly + vbInformation, "Data summary")
fx = 0
f = 0
End If
Case "7"
If shift = 0 Then
M = a
Else
fx = a + fx
f = f + 1
End If
Case "8"
B = "cbr"
Case "9"
txt1 = 1 / a
Case "16"
If shift = 0 Then
B = "xy"
Else
txt1 = 10 ^ a
End If
Case "10"
Form3.Show
Case "12"
a = a * z / 180
If shift = 0 Then
txt1 = Sin(a)
Else
txt1 = 1 / Sin(a)
End If
Case "13"
a = a * z / 180
If shift = 0 Then
txt1 = Cos(a)
Else
txt1 = 1 / Cos(a)
End If
Case "14"
a = a * z / 180
If shift = 0 Then
txt1 = Tan(a)
Else
txt1 = 1 / Tan(a)
End If
End Select
End Sub

Private Sub command3_click()
On Error GoTo handler
handler:
Resume Next
Select Case B
Case "add"
txt1.Text = a + Val(txt1.Text)
 Case "Minus"
txt1.Text = a - Val(txt1.Text)
Case "Multi"
txt1.Text = a * Val(txt1.Text)
Case "Div"
txt1 = a / Val(txt1.Text)
Case "xy"
c = txt1
txt1 = a ^ c
Case "cbr"
c = Log(a) / Log(10)
c = c / Val(txt1)
txt1 = 10 ^ c
End Select
End Sub

Private Sub command4_click()
txt1.Text = ""
End Sub


Private Sub Command5_Click()
txt1.Text = txt1.Text & "."
End Sub

Private Sub Command6_Click()
Dim quit As String
quit = MsgBox(N & ",d'you wanna give up?", vbYesNo + vbQuestion, "Quit")
If quit = vbYes Then
End
End If
End Sub

Private Sub Command7_Click()
txt1 = txt1 * -1
End Sub

Private Sub Command8_Click()
On Error GoTo handler
handler:
Resume Next
If B = "add" Then
txt1.Text = a + Val(txt1.Text) * 100
Else
If B = "Minus" Then
txt1.Text = a - Val(txt1.Text) * 100
Else
If B = "Multi" Then
txt1.Text = a * Val(txt1.Text) * 100
Else
If B = "Div" Then
txt1 = a / Val(txt1.Text) * 100

End If
End If
End If
End If
End Sub

Private Sub Command9_Click()
If shift = 1 Then
shift = 0
Command2(5).Caption = "lg"
Command2(4).Caption = "sqr"
Command2(16).Caption = "x^y"
Command2(7).Caption = "M+"
Command2(6).Caption = "MR"
Command2(12).Caption = "sin"
Command2(13).Caption = "cos"
Command2(14).Caption = "tan"
Else
If shift = 0 Then
shift = 1
Command2(5).Caption = "log"
Command2(4).Caption = "qdr"
Command2(16).Caption = "alg"
Command2(7).Caption = "D+"
Command2(6).Caption = "fx"
Command2(12).Caption = "csc"
Command2(13).Caption = "sec"
Command2(14).Caption = "cot"
End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo handler
handler:
Resume Next



N = InputBox("Please enter your name", "User Registration")
If N = "" Then
lb1.Caption = "Unregistered Version"
N = "Unregistered one"
Else
lb1.Caption = N & "'s personal Calculater"
End If

End Sub


