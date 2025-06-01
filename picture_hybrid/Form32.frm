VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   3360
   ClientLeft      =   5085
   ClientTop       =   4245
   ClientWidth     =   4740
   Icon            =   "Form32.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4740
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compression"
      Height          =   1335
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Maximum possible ratio is 50 for JPEG. Very high if more than it."
         Top             =   495
         Width           =   3000
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   350
         Left            =   210
         TabIndex        =   6
         Top             =   480
         Width           =   3065
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resolution:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Image type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Location: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filename:  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As Double
Dim b As Double

Dim w As Double

Private Sub Form_Load()
z = Form1.picture1.Width - 4
w = Form1.picture1.Height - 4
Label1 = Label1 & Form1.textfilename
Label2 = Label2 & Form1.Dir1.Path
Label2.ToolTipText = Form1.Dir1.Path
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\chimes.wav", snd_async)
Select Case Right(Form1.textfilename, 3)
Case "jpg", "JPG"
Label3 = Label3 & "JPEG interchange format"
Case "gif", "GIF"
Label3 = Label3 & "Graphic interchange format"
Case "bmp", "BMP"
Label3 = Label3 & "Windows bitmap"
Case Else
Label3 = Label3 & "image file"
End Select
b = Int(z * w / 1000)
Label4 = Label4 & b & "  Kilo Pixels"
If b > 1000 Then Label4 = "Resolution:" & Round(b / 1000, 1) & " Megapixels"

If Form1.picture1.stretch = True Then
Label4 = "Resolution: stretched"
End If
Label5 = Label5 & Form1.Label1

Select Case Right(Form1.textfilename, 3)
Case "jpg", "JPG", "jpe"
If Form1.picture1.stretch = False Then
Frame1.Visible = True
b = Int(z * w * 3)
Label8 = Int((b - Form1.Text1) / b * 100) & "% at "
w = Int(b / Form1.Text1)
Label8 = Label8 & w & " ratio"
End If
If w < 50 Then Label7.Width = w * 60 Else Label7 = "        VERY HIGH"

End Select
End Sub


