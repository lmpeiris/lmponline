VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hybrid V3.5"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "picture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1003
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   25
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   24
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7080
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   20
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Left            =   1080
      Max             =   30000
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   720
      Top             =   9720
   End
   Begin VB.HScrollBar VS 
      Height          =   135
      Left            =   4080
      TabIndex        =   17
      Top             =   9720
      Visible         =   0   'False
      Width           =   3180
   End
   Begin MCI.MMControl MMControl2 
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      ToolTipText     =   "Video controller"
      Top             =   8520
      Visible         =   0   'False
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   873
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8220
      Left            =   1080
      ScaleHeight     =   8160
      ScaleWidth      =   10800
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   10860
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   9720
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      ToolTipText     =   "Audio controller"
      Top             =   9720
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      Frames          =   30
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   624
      Left            =   120
      Max             =   1024
      SmallChange     =   32
      TabIndex        =   10
      Top             =   9360
      Width           =   12375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9255
      LargeChange     =   468
      Left            =   12600
      Max             =   768
      SmallChange     =   32
      TabIndex        =   9
      Top             =   0
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox textfilename 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3615
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   12840
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Select a file"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   12840
      TabIndex        =   1
      ToolTipText     =   "Select a directory"
      Top             =   360
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12840
      TabIndex        =   0
      ToolTipText     =   "Select a drive"
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   8760
      TabIndex        =   8
      Top             =   10200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9960
      Top             =   9600
      Width           =   765
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   23
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   10800
      TabIndex        =   22
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Raavi"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5520
      TabIndex        =   18
      Top             =   9840
      Width           =   1695
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   9600
      Width           =   3855
   End
   Begin VB.Label Label5 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   9360
      Width           =   12735
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12960
      TabIndex        =   3
      ToolTipText     =   "File size"
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13080
      TabIndex        =   7
      ToolTipText     =   "Image size"
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   " j"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   5
      ToolTipText     =   "Stretch to fit"
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label command 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   6
      ToolTipText     =   "Close"
      Top             =   9720
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   10335
      Left            =   12600
      TabIndex        =   11
      Top             =   -120
      Width           =   2535
   End
   Begin VB.Image picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   9315
      Left            =   120
      ToolTipText     =   "Double click for properties"
      Top             =   0
      Width           =   12420
   End
   Begin VB.Menu file 
      Caption         =   "&Files"
      Begin VB.Menu copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu rename 
         Caption         =   "Rename"
      End
      Begin VB.Menu Retype 
         Caption         =   "Retype"
      End
      Begin VB.Menu hyphen 
         Caption         =   "-"
      End
      Begin VB.Menu send 
         Caption         =   "Send to"
         Begin VB.Menu floppy 
            Caption         =   "floppy"
         End
         Begin VB.Menu flash 
            Caption         =   "flash drive k"
         End
         Begin VB.Menu desktop 
            Caption         =   "Desktop"
         End
      End
      Begin VB.Menu convert 
         Caption         =   "Convert BMP"
      End
      Begin VB.Menu hyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu property 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu folders 
      Caption         =   "Folder"
      Begin VB.Menu folder 
         Caption         =   "New folder"
      End
      Begin VB.Menu rfolder 
         Caption         =   "Remove folder"
      End
   End
   Begin VB.Menu settings 
      Caption         =   "Settings"
   End
   Begin VB.Menu options 
      Caption         =   "View"
      Begin VB.Menu videosize 
         Caption         =   "Video Size"
         Begin VB.Menu pal 
            Caption         =   "PAL (384x288)"
         End
         Begin VB.Menu hdtv 
            Caption         =   "HDTV (640x480)"
         End
         Begin VB.Menu dvd 
            Caption         =   "DVD (800x600)"
         End
      End
      Begin VB.Menu navigator 
         Caption         =   "Navigator"
      End
      Begin VB.Menu filetype 
         Caption         =   "File type"
         Begin VB.Menu allfiles 
            Caption         =   "All files"
         End
         Begin VB.Menu jpegs 
            Caption         =   "All pictures"
         End
         Begin VB.Menu jpg2 
            Caption         =   "jpegs"
         End
         Begin VB.Menu bitmaps 
            Caption         =   "Bitmaps"
         End
         Begin VB.Menu audio 
            Caption         =   "Audio"
         End
         Begin VB.Menu video 
            Caption         =   "Video"
         End
         Begin VB.Menu customtype 
            Caption         =   "custom"
         End
      End
      Begin VB.Menu help 
         Caption         =   "help"
         Begin VB.Menu subhelp 
            Caption         =   "help"
         End
         Begin VB.Menu hybrid 
            Caption         =   "about Hybrid"
         End
      End
      Begin VB.Menu autoplay 
         Caption         =   "Autoplay"
         Checked         =   -1  'True
      End
      Begin VB.Menu hidden2 
         Caption         =   "Show hidden"
      End
      Begin VB.Menu slide 
         Caption         =   "Slideshow"
      End
   End
   Begin VB.Menu effects 
      Caption         =   "Extras"
      Begin VB.Menu Playlists 
         Caption         =   "Playlists"
         Begin VB.Menu load 
            Caption         =   "Load "
         End
         Begin VB.Menu save 
            Caption         =   "Save"
         End
         Begin VB.Menu clear 
            Caption         =   "Clear"
         End
      End
      Begin VB.Menu playcd 
         Caption         =   "play audioCD"
      End
      Begin VB.Menu wav 
         Caption         =   "stop wav"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As Double
Dim d As Double
Dim f As Integer
Dim navig As Integer
Dim pathchange As Boolean
Dim janaka As String
Dim mode As Integer
Dim c As String
Dim tracktime As String * 10
Dim tracktime2 As String * 12
Dim copy1 As String
Dim copy2 As String
Dim malshan As String
Dim OFS As OFSTRUCT
Public hfile As Long
Public filename As String


Private Sub allfiles_Click()
File1.Pattern = "*.*"
End Sub



Private Sub audio_Click()
File1.Pattern = "*.mp3;*.MP3;*.wma;*.WMA;*.wav;*.WAV;*.MID*;*.mid*"
End Sub

Private Sub autoplay_Click()
autoplay.Checked = Not autoplay.Checked
End Sub

Private Sub bitmaps_Click()
File1.Pattern = "*.bmp*;*.BMP*"
End Sub

Private Sub clear_Click()
Combo1.clear
End Sub

Private Sub Combo1_Click()

VS.Visible = True
Label6 = "Playlist-Track " & Combo1.ListIndex + 1
MMControl1.command = "close"
MMControl1.Visible = True
MMControl1.filename = Combo1.Text
MMControl1.command = "open"
MMControl1.command = "play"
b = MMControl1.TrackLength
VS.Max = b / 1000
d = b / 60000
b = d - Int(d)
b = Int(b * 60)
tracktime = Int(d) & ":" & b
File1.SetFocus

End Sub

Private Sub Combo2_Click()
Dir1.Path = Combo2.Text
Drive1.Drive = Left(Dir1.Path, 1)
End Sub

Private Sub Command1_Click()
navig = navig - 1
If navig > -Combo2.ListCount Then Command1.Enabled = True Else Command1.Enabled = False
pathchange = True
Dir1.Path = Combo2.List(Combo2.ListCount + navig - 1)
Drive1.Drive = Left(Dir1.Path, 1)
pathchange = False
If Command2.Enabled = False Then Command2.Enabled = True
End Sub

Private Sub Command2_Click()
navig = navig + 1
If navig < 0 Then Command2.Enabled = True Else Command2.Enabled = False
pathchange = True
Dir1.Path = Combo2.List(Combo2.ListCount + navig - 1)
Drive1.Drive = Left(Dir1.Path, 1)
pathchange = False
If Command1.Enabled = False Then Command1.Enabled = True
End Sub

Private Sub convert_Click()
If Right(Dir1.Path, 1) <> "\" Then malshan = Dir1.Path & "\" Else malshan = Dir1.Path
a = InputBox("Enter a name for the new file", "Convert")
If a <> "" Then SavePicture picture1.Picture, malshan & a & ".bmp"
File1.Refresh
End Sub



Private Sub Command_Click()
Dim cl As String
cl = MsgBox("Do you really want to quit?", vbQuestion + vbYesNo, "Quit")
If cl = vbYes Then
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\tada.wav", snd_async)
End
End If
End Sub





Private Sub customtype_Click()
Dim custom As String
custom = InputBox("Type the extension of the filetype you want to filter via browser" & Chr$(13) & "Ex: bmp for bitmap files", "Custom filter", "wmv")
If custom <> "" Then File1.Pattern = "*." & custom & "*"
End Sub

Private Sub desktop_Click()
FileCopy filename, "C:\Documents and Settings\All Users\Desktop" & "\" & File1.filename
End Sub

Private Sub dvd_Click()
Picture2.Height = 548
Picture2.Left = 72
Picture2.Top = 8
Picture2.Width = 724
MMControl2.Height = 33
MMControl2.Left = 72
MMControl2.Top = 568
MMControl2.Width = 724
HS.Height = 17
HS.Left = 72
HS.Top = 552
HS.Width = 721
End Sub

Private Sub File1_PathChange()
Label3 = "Total " & File1.ListCount & " files"
Label1 = ""
End Sub

Private Sub File1_PatternChange()
Label3 = "Total " & File1.ListCount & " files"
Label1 = ""
End Sub

Private Sub flash_Click()
On Error GoTo handler3:
start2:
FileCopy filename, Right(flash.Caption, 1) & ":\" & File1.filename
Exit Sub
handler3:
a = MsgBox("Please insert a writable flash\pen drive to a USB slot and continue. If problem continues click settings and go on.", vbCritical + vbRetryCancel, Err.Description)
If a = vbRetry Then GoTo start2 Else Exit Sub
End Sub

Private Sub folder_Click()
a = InputBox("Enter name for the new folder", "New folder")
If Right(Dir1.Path, 1) <> "\" Then malshan = Dir1.Path & "\" Else malshan = Dir1.Path
If a <> "" Then MkDir malshan & a
Dir1.Refresh
End Sub



Private Sub Form_Load()
Open App.Path & "\files\settings.ini" For Input As #5
Input #5, malshan
flash.Caption = "flash drive " & malshan
Input #5, malshan
Timer1.Interval = Val(malshan)
Input #5, malshan
Drive1.Drive = Left(malshan, 3)
Dir1.Path = malshan
Input #5, malshan
If malshan = 1 Then i = True Else i = False
Close #5
If Screen.Height <> 11520 Then MsgBox "Please set resolution to 1024 x 768 for full performance."
VScroll1.Visible = False
HScroll1.Visible = False
picture1 = LoadPicture(App.Path & "\files\theme.jpg")
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\Windows XP Startup.wav", snd_async)
navig = 0
End Sub

Private Sub hdtv_Click()
Picture2.Height = 484
Picture2.Left = 112
Picture2.Top = 72
Picture2.Width = 644
MMControl2.Height = 33
MMControl2.Left = 112
MMControl2.Top = 560
MMControl2.Width = 644
HS.Height = 9
HS.Left = 112
HS.Top = 552
HS.Width = 641
End Sub

Private Sub hidden2_Click()
File1.Hidden = Not File1.Hidden
hidden2.Checked = Not hidden2.Checked
End Sub

Private Sub HScroll1_Change()
picture1.Left = 0 - HScroll1
End Sub



Private Sub hybrid_Click()
MsgBox " Hybrid version 3.5 is built by Lasika Malshan Peiris(Sir LMP-email to lmpeiris@gmail.com\ Mobile- +94-0775525110)" & Chr$(13) & " Copyrighted 2003-2007 by LMP Online inc." & Chr$(13) & " Version " & App.Major & "." & App.Minor & "." & App.Revision, , "Hybrid V3.0- Hybridization of pictureviewing"
End Sub

Private Sub jpegs_Click()
File1.Pattern = "*.jpg;*.bmp;*.gif;*.ico;*.wmf;*.jpe;*.JPG*;*.BMP*;*.gif*"
End Sub

Private Sub jpg2_Click()
File1.Pattern = "*.jpg*;*.JPG*;*.jpe*"
End Sub






Private Sub Label2_Click()
If Label2.BorderStyle = 0 Then
picture1.Top = 0
picture1.Left = 8
Label2.BorderStyle = 1

picture1.Width = 828
picture1.Height = 621
VScroll1.Visible = False
HScroll1.Visible = False
Else
Label2.BorderStyle = 0

End If
End Sub












Private Sub cut_Click()
mode = 2
copy1 = filename
copy2 = File1.filename
End Sub

Private Sub delete_Click()
On Error GoTo handler4
a = "Are you sure that you really want to delete " & filename & "?"
a = MsgBox(a, vbExclamation + vbYesNo, "Confirm file delete")
If a = vbYes Then
DeleteFile filename
File1.Refresh
textfilename = ""
End If
Exit Sub
handler4:
a = MsgBox("Unable to delete " & filename & Chr$(13) & Err.Description, vbCritical + vbOKOnly, "Can't delete the file")
Exit Sub
End Sub

Private Sub file_click()
If copy1 = "" Then paste.Enabled = False Else paste.Enabled = True
If textfilename = "" Then
copy.Enabled = False
send.Enabled = False
delete.Enabled = False
cut.Enabled = False
rename.Enabled = False
Retype.Enabled = False
Else
copy.Enabled = True
send.Enabled = True
delete.Enabled = True
cut.Enabled = True
rename.Enabled = True
Retype.Enabled = True
End If
If Drive1.Drive = "a:" Then floppy.Enabled = False
If textfilename = "" Then
convert.Enabled = False
Else
convert.Enabled = True

End If
Select Case Right(textfilename, 3)
Case "jpg", "JPG", "jpe"
convert.Enabled = True
Case "GIF", "gif"
convert.Enabled = True
Case Else
convert.Enabled = False

End Select
End Sub



Private Sub copy_Click()
mode = 1
copy1 = filename
copy2 = File1.filename
End Sub

Private Sub floppy_Click()
On Error GoTo handler3:
start2:
FileCopy filename, "a:\" & File1.filename
Exit Sub
handler3:
a = MsgBox("Please insert a writable floppy disk to the floppy drive and continue.", vbCritical + vbRetryCancel, Err.Description)
If a = vbRetry Then GoTo start2 Else Exit Sub
End Sub

Private Sub Dir1_Change()
property.Enabled = False
File1.Path = Dir1.Path
textfilename = ""
f = 0
If pathchange = False Then Combo2.AddItem Dir1.Path
Command1.Enabled = True
Label8 = "Folder"
Label9 = Dir1.Path
Image1.Picture = LoadPicture(App.Path & "\icons\folder.bmp")
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\start.wav", snd_async)
End Sub

Private Sub Drive1_Change()
On Error GoTo hand2:
property.Enabled = False
Dir1.Path = Drive1.Drive
File1.Path = Drive1.Drive
Dim disktype As Long
disktype = GetDriveType(Left(Drive1.Drive, 1) & ":\")
Select Case disktype
Case 0
Label8 = "Unidentified drive"
Case 1
Label8 = "No drive"
Case 2
If Left(Drive1.Drive, 1) <> "a" Then
Label8 = "Flash drive"
Image1.Picture = LoadPicture(App.Path & "\icons\flash.jpg")
Else
Label8 = "Floppy drive"
End If
Case 3
Label8 = "NTFS/FAT32 drive"
Image1.Picture = LoadPicture(App.Path & "\icons\hard.jpg")
Case 5
Label8 = "CD/DVD drive"
Image1.Picture = LoadPicture(App.Path & "\icons\cd.jpg")
Case Else
Label8 = "Unidentified drive"
End Select

Dim lpSectorsPerCluster As Long
Dim lpBytesPerSector As Long
Dim lpNumberOfFreeClusters As Long
Dim lpTotalNumberOfClusters As Long
Dim dbyte As Double
disktype = GetDiskFreeSpace(Left(Drive1.Drive, 1) & ":\", lpSectorsPerCluster, lpBytesPerSector, lpNumberOfFreeClusters, lpTotalNumberOfClusters)
'free MB
dbyte = lpNumberOfFreeClusters / 1024 ^ 2
Label9 = Int(lpSectorsPerCluster * lpBytesPerSector * dbyte)
'total MB
dbyte = lpTotalNumberOfClusters / 1024 ^ 2
Label9 = Label9 & " / " & Int(lpSectorsPerCluster * lpBytesPerSector * dbyte) & " MB free"

Exit Sub
hand2:
MsgBox Err.Description, vbCritical, "Error accesing device"
Drive1.Drive = "c:\"
Exit Sub
End Sub

Private Sub File1_Click()
property.Enabled = True
Label8 = ""
Label9 = ""
Image1.Picture = LoadPicture("")
textfilename.ToolTipText = File1.filename
MMControl2.command = "stop"
HS.Visible = False
MMControl2.command = "close"
MMControl2.Visible = False
Picture2.Visible = False
HScroll1.Visible = False
VScroll1.Visible = False
Label3 = ""
picture1 = LoadPicture("")
picture1.Visible = False
On Error GoTo handler
textfilename = File1.filename
If Right(Dir1.Path, 1) = "\" Then
filename = Dir1.Path & textfilename
Else
filename = Dir1.Path & "\" & textfilename
End If

Dim attr As Integer
attr = GetFileAttributes(filename)

Select Case attr
Case 128
Label9 = "r w _ _"
Case 1
Label9 = "r _ _ _"
Case 2
Label9 = "r w _ h"
Case 32
Label9 = "r w x _"
Case 33
Label9 = "r _ x _"
Case 35
Label9 = "r _ x h"
Case 34
Label9 = "r w x h"
Case 3
Label9 = "r _ _ h"
Case Else
Label9 = "_ _ _ _"
End Select

Open filename For Random As #1
Text1 = LOF(1)
If LOF(1) > 1024 Then
a = LOF(1) / 1024
b = Val(a)
Label1.Caption = Round(b) & " kB"
Else
Label1.Caption = LOF(1) & " Bytes"
End If
If LOF(1) > 1048576 Then
a = LOF(1) / 1048576
b = Val(a)
Label1 = Round(b, 1) & " MB"
End If
Close #1

Select Case Right(filename, 3)
Case "JPG", "bmp", "gif", "BMP", "GIF", "ico", "ICO", "jpg", "jpe", "wmf", "JPE", "WMF"
Label8 = "Image file"
Image1.Picture = LoadPicture(App.Path & "\icons\image.jpg")
picture1.Visible = True
Form1.Caption = "Hybrid V3.5-" & textfilename
picture1.stretch = False
VScroll1.Left = 840
HScroll1.Top = 624
picture1.Top = 0
picture1.Left = 0
HScroll1 = 0
VScroll1 = 0
Form1.MousePointer = 11
picture1 = LoadPicture(filename)
a = picture1.Height - 4
b = picture1.Width - 4
Label3 = b & " x " & a & "  Pixels"
Form1.MousePointer = 0
If Label2.BorderStyle = 1 Then GoTo stretch Else GoTo scroll


Case "wav", "WAV"
Label8 = "Raw sound file"
If autoplay.Checked = True Then
Dim lmp As String
lmp = sndPlaySound(filename, snd_async)
End If
GoTo finish

Case "AVI", "avi", "mpg", "MPG", "wmv", "WMV"
Label8 = "Video file"
Image1.Picture = LoadPicture(App.Path & "\icons\video.jpg")
If autoplay.Checked = True Then
Picture2.Visible = True
Picture2.ToolTipText = textfilename
Form1.Caption = "Hybrid V3.5-" & textfilename
Picture2.Cls
MMControl2.command = "close"
MMControl2.Visible = True
MMControl2.hWndDisplay = Picture2.hWnd
MMControl2.filename = filename
MMControl2.command = "open"
tracktime2 = MMControl2.TrackLength
HS.Visible = True
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\Windows XP Hardware Insert.wav", snd_async)
If tracktime2 > 32000 Then HS.Max = MMControl2.TrackLength / 1000 Else HS.Max = MMControl2.TrackLength
MMControl2.command = "play"
End If
If Right(textfilename, 3) = "wmv" Or Right(textfilename, 3) = "WMV" Then
b = MMControl2.TrackLength / 1000
Open filename For Random As #3
d = LOF(3) / 1024
Close #3
Label3 = Int((d / b) * 8.52) & " Kbps"
End If

GoTo finish

Case "MP3", "mp3", "WMA", "wma", "mid", "MID"
Label8 = "Audio file"
Image1.Picture = LoadPicture(App.Path & "\icons\audio.jpg")
Combo1.AddItem filename
Combo1.Visible = True
If autoplay.Checked = True Then
VS.Visible = True
Label6 = textfilename
MMControl1.command = "close"
MMControl1.Visible = True
MMControl1.filename = filename
MMControl1.command = "open"
MMControl1.command = "play"
b = MMControl1.TrackLength
VS.Max = b / 1000
d = b / 60000
b = d - Int(d)
b = Int(b * 60)
tracktime = Int(d) & ":" & b
b = MMControl1.TrackLength / 1000
Open filename For Random As #2
d = LOF(2) / 1024
Close #2
Label3 = Int((d / b) * 8.2) & " Kbps"
Label6 = Label6 & " (" & Label3 & ")"
End If
GoTo finish

Case "plt"
Label8 = "Hybrid playlist file"
Image1.Picture = LoadPicture(App.Path & "\icons\audio.jpg")
Combo1.Visible = True
Combo1.clear
Open filename For Input As #3
GoTo down
down:
If EOF(3) = False Then
Input #3, janaka
Combo1.AddItem janaka
GoTo down
Else
Close #3
Exit Sub
End If
Close #3
GoTo finish

Case "exe", "EXE"
Label8 = "Application"
Image1.Picture = LoadPicture(App.Path & "\icons\system.jpg")
Case "txt", "TXT"
Label8 = "Text file"
Image1.Picture = LoadPicture(App.Path & "\icons\text.jpg")
Case "dll", "DLL"
Label8 = "Direct Link Library"
Image1.Picture = LoadPicture(App.Path & "\icons\exe.jpg")
Case "doc", "DOC"
Label8 = "MS Word document"
Image1.Picture = LoadPicture(App.Path & "\icons\doc.jpg")
Case "ppt", "PPT", "pps", "PPS"
Label8 = "Powerpoint slide"
Image1.Picture = LoadPicture(App.Path & "\icons\ppt.jpg")
Case "ini", "INI", "htt", "HTT", "inf"
Label8 = "Configuration file"
Image1.Picture = LoadPicture(App.Path & "\icons\exe.jpg")
Case "htm", "HTM", "tml", "TML"
Label8 = "HTML document"
Image1.Picture = LoadPicture(App.Path & "\icons\html.jpg")
Case "zip", "ZIP", "rar", "RAR", ".gz", "tar"
Label8 = "Compressed archeive"
Image1.Picture = LoadPicture(App.Path & "\icons\zip.jpg")
Case "mp4", "MP4", "vob", "VOB", "ivx", "ivd"
Label8 = "Video files"
Image1.Picture = LoadPicture(App.Path & "\icons\video.jpg")
Case "ogg", "OGG"
Label8 = "Audio files"
Image1.Picture = LoadPicture(App.Path & "\icons\audio.jpg")
Case "pdf", "PDF"
Label8 = "Acrobat document"
Image1.Picture = LoadPicture(App.Path & "\icons\pdf.jpg")
Case "sys", "SYS"
Label8 = "System file"
Image1.Picture = LoadPicture(App.Path & "\icons\exe.jpg")
Case Else

hfile = OpenFile(filename, OFS, &H2)
Call CloseHandle(hfile)
End Select

scroll:
If picture1.Height > 621 Then
VScroll1.Visible = True
VScroll1.Max = picture1.Height - 621
If picture1.Width < 828 Then
VScroll1.Left = 8 + picture1.Width
End If
End If
If picture1.Width > 828 Then
HScroll1.Visible = True
HScroll1.Max = picture1.Width - 828
If picture1.Height < 621 Then
HScroll1.Top = picture1.Height
End If
End If
If a < 621 And b < 828 Then
picture1.Top = Int(621 - a) / 2
picture1.Left = 8 + Int(828 - b) / 2
End If
GoTo finish

stretch:
If a > b Then
d = b / a
picture1.Width = Int(d * 600)
picture1.Height = 600
Else
d = a / b
picture1.Height = Int(d * 800)
picture1.Width = 800
End If
picture1.stretch = True
GoTo finish

handler:
Resume Next
finish:
End Sub


Private Sub load_Click()
Dir1.Path = App.Path & "\files\"
File1.Path = App.Path & "\files\"
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\chimes.wav", snd_async)
End Sub

Private Sub navigator_Click()
navigator.Checked = Not navigator.Checked
Combo2.Visible = Not Combo2.Visible
If i = True Then janaka = sndPlaySound(App.Path & "\sounds\chimes.wav", snd_async)
End Sub

Private Sub pal_Click()
Picture2.Height = 292
Picture2.Left = 292
Picture2.Top = 184
Picture2.Width = 356
MMControl2.Height = 33
MMControl2.Left = 292
MMControl2.Top = 488
MMControl2.Width = 356
HS.Height = 9
HS.Left = 292
HS.Top = 476
HS.Width = 353
End Sub

Private Sub paste_Click()
On Error GoTo hand4:
Select Case mode
Case 1
filename = "Are you sure you want to copy " & copy1 & " to " & Dir1.Path & "?"
filename = MsgBox(filename, vbExclamation + vbYesNo, "Confirm file copying")
If filename = vbYes Then
If Right(Dir1.Path, 1) <> "\" Then malshan = Dir1.Path & "\" Else malshan = Dir1.Path
FileCopy copy1, malshan & copy2
copy1 = ""
File1.Refresh
textfilename = ""
Exit Sub

End If
Case 2
filename = "Are you sure you want to move " & copy1 & " to " & Dir1.Path & "?"
filename = MsgBox(filename, vbExclamation + vbYesNo, "Confirm file transform")
If filename = vbYes Then
If Right(Dir1.Path, 1) <> "\" Then malshan = Dir1.Path & "\" Else malshan = Dir1.Path
FileCopy copy1, malshan & copy2
Kill copy1
copy1 = ""
File1.Refresh
End If
End Select
Exit Sub
hand4:
a = MsgBox(Err.Description & ".Unable to transfer", vbCritical + vbOKOnly, "Error...pasting " & copy1)
End Sub



Private Sub picture1_DblClick()
If picture1.Picture <> 0 Then Form3.Show
End Sub


Private Sub playcd_Click()
If playcd.Checked = False Then
playcd.Checked = True
MMControl1.Visible = True
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.EjectVisible = True
MMControl1.NextVisible = True
MMControl1.PrevVisible = True
MMControl1.DeviceType = "CDAudio"
MMControl1.command = "Open"
tracktime = MMControl1.Length
Else
MMControl1.command = "stop"
MMControl1.command = "close"
MMControl1.Visible = False
playcd.Checked = False
End If
End Sub



Private Sub property_Click()
Form2.Show
End Sub

Private Sub rename_Click()
On Error GoTo handler6:
If Dir1.Path <> "\" Then malshan = Dir1.Path & "\" Else malshan = Dir1.Path
a = InputBox("Enter a new name for the file", "Confirm file rename")
If a <> "" Then
Name filename As malshan & a & Right(File1.filename, 4)
File1.Refresh
textfilename = ""
End If
Exit Sub

handler6:
MsgBox " Please give a different name.", vbCritical + vbOKOnly, Err.Description
End Sub








Private Sub Retype_Click()
On Error GoTo handler6
a = InputBox("Enter three letters or numbers without any other signs. Ex: txt", "Change extension", "mpg")
If a <> "" Then
c = Mid(filename, 1, Len(filename) - 3)
Name filename As c & a
File1.Refresh
textfilename = ""
End If
Exit Sub
handler6:
MsgBox " Please give a different name.", vbCritical + vbOKOnly, Err.Description
End Sub

Private Sub rfolder_Click()
malshan = "Are you sure you want to delete the folder " & Dir1.Path & "?" & Chr$(13) & "Includes " & File1.ListCount & " files."
a = MsgBox(malshan, vbExclamation + vbYesNo, "Folder remove")
On Error GoTo handler9:
If a = vbYes Then
c = Dir1.Path & "\*.*"
If File1.ListCount <> 0 Then Kill c
RmDir Dir1.Path & "\"
End If
Exit Sub
handler9:
a = MsgBox("Can't remove " & c & Chr$(13) & Err.Description, vbCritical + vbOKOnly, "Error removing directory")
End Sub


Private Sub save_Click()
malshan = InputBox("Please type a filename for the the new playlist file.", "What will you call it?", "History")
If malshan <> "" Then
Open App.Path & "\files\" & malshan & ".plt" For Output As #4
For n = 1 To Combo1.ListCount - 1
Print #4, Combo1.List(n)
Next n
Close #4
End If
End Sub

Private Sub settings_Click()
Form31.Show
End Sub

Private Sub slide_Click()
slide.Checked = Not slide.Checked
f = 0
End Sub

Private Sub subhelp_Click()
On Error GoTo down
File1.Path = App.Path & "\help"
Dir1.Path = App.Path & "\help"
File1.SetFocus
File1.Selected(0) = True
Exit Sub
down:
MsgBox "Cannot find the help files supposed to be located in " & App.Path & "\help", vbCritical + vbOKOnly, "Cannot find help files"
End Sub

Private Sub textfilename_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
PopupMenu file
End If
End Sub



Private Sub Timer1_Timer()
If slide.Checked = True Then

If Timer1.Interval Then
If f + 1 < File1.ListCount Then
f = File1.ListIndex + 1
File1.Selected(f - 1) = False
File1.Selected(f) = True
End If
End If
End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo hand:
If MMControl1.Visible = True Then
If playcd.Checked = False Then
If Timer2.Interval Then
VS = MMControl1.Position / 1000
b = MMControl1.Position
d = b / 60000
b = d - Int(d)
b = Int(b * 60)
Label7 = Int(d) & ":" & b & "/" & tracktime
End If
End If
End If
If MMControl2.Visible = True Then
If Timer2.Interval Then
If tracktime2 > 32000 Then
HS = MMControl2.Position / 1000
Else
HS = MMControl2.Position
End If
End If
End If
Exit Sub
hand:
MMControl2.command = "stop"
MMControl2.command = "close"
End Sub

Private Sub video_Click()
File1.Pattern = "*.mpg;*.MPG;*.wmv;*.WMV;*.avi;*.AVI"
End Sub

Private Sub VScroll1_Change()
picture1.Top = 0 - VScroll1
End Sub


Private Sub wav_Click()
Dim lmp As String
Label3 = "stopped " & textfilename
lmp = sndPlaySound(App.Path & "\sounds\tada.wav", snd_async)
End Sub
