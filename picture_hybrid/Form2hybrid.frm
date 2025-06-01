VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "General file properties"
   ClientHeight    =   1155
   ClientLeft      =   8775
   ClientTop       =   5775
   ClientWidth     =   5085
   LinkTopic       =   "Form2"
   ScaleHeight     =   1155
   ScaleWidth      =   5085
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form2hybrid.frx":0000
      Left            =   240
      List            =   "Form2hybrid.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "File Attributes- (r-read,w-write,x-archieve,h-hidden}"
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
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
Dim damn As Long
Dim attr As Long
Select Case Combo1.Text
Case "r w _ _"
attr = 128
Case "r _ _ _"
attr = 1
Case "r w _ h"
attr = 2
Case "r w x _"
attr = 32
Case "r _ x _"
attr = 33
Case "r _ x h"
attr = 35
Case "r w x h"
attr = 34
Case "r _ _ h"
attr = 3
End Select
If attr <> 0 Then damn = SetFileAttributes(Form1.filename, attr)
Form1.Label9 = Combo1.Text
attr = 0
Form1.File1.SetFocus
Unload Me
End Sub

Private Sub Form_Load()
Combo1.AddItem "r w _ _"
Combo1.AddItem "r _ _ _"
Combo1.AddItem "r w _ h"
Combo1.AddItem "r w x _"
Combo1.AddItem "r _ x _"
Combo1.AddItem "r _ x h"
Combo1.AddItem "r w x h"
Combo1.AddItem "r _ _ h"
Combo1.AddItem "_ _ _ _"
Combo1.Text = Form1.Label9
End Sub


