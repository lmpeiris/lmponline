VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form31 
   Caption         =   "Settings"
   ClientHeight    =   3930
   ClientLeft      =   4815
   ClientTop       =   3225
   ClientWidth     =   4650
   Icon            =   "Form31.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4650
   Begin VB.Frame Frame1 
      Caption         =   "Activations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Play program menu sounds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Value           =   5
      Max             =   30
      Min             =   2
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Time interval between two slides in a slide show(seconds):"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Default USB Drive letter:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Start up directory:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As String
Dim oshadee As String


Private Sub Command2_Click()
oshadee = InputBox("Enter the letter which is allocated to a flash drive when inserted to your system.", "USB flash/pen drive identification")
If Len(oshadee) = 1 And Val(oshadee) = 0 Then Label2 = "Default USB Drive letter: " & oshadee
End Sub

Private Sub Command3_Click()
Form1.flash.Caption = "flash drive" & Right(Label2, 1)
Form1.Timer1.Interval = Val(Label4) * 1000
If Check1.Value = 1 Then i = True Else i = False
Open App.Path & "\files\settings.ini" For Output As #5
Print #5, Right(Label2, 1)
Print #5, Form1.Timer1.Interval
Print #5, Text1
Print #5, Check1.Value
Close #5
Unload Me
End Sub

Private Sub Form_Load()
Label4 = Form1.Timer1.Interval / 1000
UpDown1.Value = Val(Label4)
Label2 = "Default USB Drive letter: " & Right(Form1.flash.Caption, 1)
Open App.Path & "\files\settings.ini" For Input As #5
Input #5, z
Input #5, z
Input #5, z
Text1 = z
Input #5, z
Check1.Value = z
Close #5
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\chimes.wav", snd_async)
End Sub



Private Sub UpDown1_Change()
Label4 = UpDown1.Value
End Sub
