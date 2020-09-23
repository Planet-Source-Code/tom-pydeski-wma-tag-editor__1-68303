VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtStatus 
      Height          =   1020
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   6465
      Width           =   9615
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H00800000&
      Height          =   250
      Left            =   735
      TabIndex        =   8
      Text            =   "Editing"
      Top             =   315
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5265
      Left            =   45
      TabIndex        =   7
      Top             =   420
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   9287
      _Version        =   393216
      Rows            =   20
      ScrollTrack     =   -1  'True
   End
   Begin ComctlLib.TreeView tvHeader 
      Height          =   3060
      Left            =   5040
      TabIndex        =   6
      Top             =   3500
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   5398
      _Version        =   327682
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8190
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   5670
      Width           =   9615
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10800
      TabIndex        =   1
      Top             =   50
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   10695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9765
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest.frx":071E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest.frx":0A38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mBar 
         Caption         =   "-"
      End
      Begin VB.Menu mNote 
         Caption         =   "Edit File with MDINote"
      End
      Begin VB.Menu mBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mPlay 
         Caption         =   "Play File with MediaPlayer"
      End
      Begin VB.Menu mExplore 
         Caption         =   "Launch Explorer"
      End
      Begin VB.Menu mseps 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu medit 
      Caption         =   "&Edit"
      Begin VB.Menu mcut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mpaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mpastespec 
         Caption         =   "Paste &Special"
      End
      Begin VB.Menu mBarE 
         Caption         =   "-"
      End
      Begin VB.Menu mCase 
         Caption         =   "Change C&ase"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mBarE1 
         Caption         =   "-"
      End
      Begin VB.Menu mCopyName 
         Caption         =   "Copy File&Name to Clipboard"
      End
      Begin VB.Menu mCopyTree 
         Caption         =   "Copy Treeview Data to Clipboard"
      End
      Begin VB.Menu mundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'submitted by Tom Pydeski
'I saw a lot of examples for editing mp3 tags, but not much to read wma files.
'there was one by Somenon at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61254&lngWId=1
'Although the original author's code worked, I wanted a way to read the file
'without parsing through each charater looking for a certain string.
'(I kept his original routines as reference)
'
'I utilized the class structure from InfoTag, which read WMA files, but did it in a way
'that would not allow writing back to the file.
'So I initially tried to dig into the file and try to read it in blocks.
'It wasn't long before I realized the structure was way more complicated than I had
'originally thought.  I did some digging and found the attached document
'"Advanced Systems Format (ASF) Specification" from Microsoft.
'Using this as a guide, I built the structures neccessary for each header object.
'I spent many weeks developing this and my wife hated that I was always on the 'puter,
'but I wanted to finish this.  It will read and write the basic tags and I put in a
'treeview to display the entire file structure by object.
'I'd like to add another flexgrid and read multiple files, but that's for later.
'
'The other thing that I had hoped to accomplish was this:
'When using my GotRadio submission, I found that temp files were created with the filename of
'the songs played.  These were located in the temporary internet directory.
'The structure of these files is different than the structure of wma's that I had burned.
'I was hoping to be able to learn how to modify the temp files to allow them to be played
'by media player, but try as i might, when I converted the temp. file to a wma with the set
'structure, media player would not play it.
'I'm attaching one of those temp files before and after modifying it.
'maybe someone with more knowledge can find the errors of my way
'
Dim WMATag As TagReader.InfoTag
Dim TagSuccess As Boolean
Dim fName$
Dim eTitle$
Dim eMess$
Dim mError As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'Listbox API constant
Const LB_ITEMFROMPOINT = &H1A9
Const LB_SETTOPINDEX = &H197
Const LB_FINDSTRING = &H18F
Const LB_SELECTSTRING = &H18C
Const LB_SELITEMRANGEEX = &H183
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Dim lRet As Long
Dim Ignore As Byte
Dim Changing As Byte
Dim IgnoreClick As Boolean
Dim IgnoreTextChange As Boolean
Dim gText As String
'for moving stuff
Private Type POINTAPI
    X As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim NewClip As String
Dim pVal As Integer
Dim LastpVal As Integer
Dim chN As Integer
Dim chIn$
Dim i As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then
    'we went to move the txtedit if it has focus
    Beep
End If
End Sub

Private Sub Form_Load()
'initialize the flex grid to display tag info
With Grid1
    .ColWidth(0) = 1800
    .ColWidth(1) = 7000
    .Rows = 20
    .Cols = 2
    .TextMatrix(0, 0) = "Tag Name"
    .TextMatrix(0, 1) = "Tag Value"
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(1, 0) = "Title "
    .TextMatrix(2, 0) = "Artist "
    .TextMatrix(3, 0) = "Album "
    .TextMatrix(4, 0) = "AlbumArtist "
    .TextMatrix(5, 0) = "Composer "
    .TextMatrix(6, 0) = "Publisher "
    .TextMatrix(7, 0) = "FileSize "
    .TextMatrix(8, 0) = "TrackNumber "
    .TextMatrix(9, 0) = "Year "
    .TextMatrix(10, 0) = "Bitrate "
    .TextMatrix(11, 0) = "Frequency "
    .TextMatrix(12, 0) = "IsCopyright "
    .TextMatrix(13, 0) = "Copyright "
    .TextMatrix(14, 0) = "IsLicensed "
    .TextMatrix(15, 0) = "IsProtected "
    .TextMatrix(16, 0) = "Rating "
    .TextMatrix(17, 0) = "Duration "
    .TextMatrix(18, 0) = "Mode "
    .TextMatrix(19, 0) = "Comments "
    .Height = 5265
End With
txtFile.Text = GetSetting(App.EXEName, "Settings", "LastFile", txtFile.Text)
Dir1.Path = Left$(txtFile.Text, InStrRev(txtFile.Text, "\"))
File1.Path = Dir1.Path
File1.Refresh
txtInfo.SelText = "    File Length   " & " |  " & "ASF_QSize" & " | " & "  CD_QSize" & " | " & " Reserved " & " | " & "ExtCont_QSize" & vbCrLf
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
cmdSelect.Left = (Me.ScaleWidth - cmdSelect.Width) - 50
txtFile.Width = cmdSelect.Left - 100
'resize everything to fit nice
File1.Top = Grid1.Top + 50
Drive1.Top = File1.Top
Dir1.Top = Drive1.Top + Drive1.Height + 25
File1.Left = (Me.ScaleWidth - File1.Width) - 100
Drive1.Left = File1.Left - Dir1.Width - 25
Dir1.Left = Drive1.Left
'Dir1.Height = File1.Height - Drive1.Height
With tvHeader
    .Left = Dir1.Left
    .Top = File1.Top + File1.Height + 50
    .Width = Me.ScaleWidth - .Left - 50
    .Height = Me.ScaleHeight - .Top - 100
End With
'
DoEvents
With Grid1
    .Move 50, .Top, Dir1.Left - 100 ', (Me.ScaleHeight - txtInfo.Height) - .Top - 100 ', (Me.ScaleHeight - .Top) - 100
End With
With txtInfo
    .Move 50, (Grid1.Top + Grid1.Height) + 50, Dir1.Left - 100, ((Me.ScaleHeight - .Top) - 100) / 2
End With
With txtStatus
    .Move 50, (txtInfo.Top + txtInfo.Height) + 50, Dir1.Left - 100, txtInfo.Height
End With
End Sub

Private Sub Grid1_Click()
Changing = 0
Rowp = Grid1.Row
Colp = Grid1.Col
Scratch$ = Grid1.TextMatrix(Rowp, Colp)
End Sub

Private Sub Grid1_DblClick()
On Error GoTo Oops
Dim tX As Integer
Dim tY As Integer
'we don't need to sort, but i threw this in because i may add a grid to list many songs
If Grid1.Row = 1 And Grid1.RowSel = Grid1.Rows - 1 Then
    Grid1.RowSel = Grid1.Rows - 2
    Grid1.Sort = flexSortGenericAscending 'flexSortNumericAscending
    Grid1.Row = 1
    Grid1.RowSel = 1
    Grid1.TopRow = 1
    Exit Sub
End If
Changing = 0
Rowp = Grid1.Row
Colp = Grid1.Col
If Grid1.Col < 7 Then
    Scratch$ = Grid1.TextMatrix(Rowp, Colp)
Else
    Scratch$ = Grid1.TextMatrix(Rowp, 1)
End If
'Grid1.SetFocus
Grid1.Tag = Scratch$
Set_TextBox
Me.txtEdit.SetFocus
tX = (txtEdit.Left + txtEdit.Width - 1) \ Screen.TwipsPerPixelX
tY = (txtEdit.Top + 700) \ Screen.TwipsPerPixelX
Dirty = True
Refresh
DoEvents
GoTo Exit_Grid1_DblClick
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_DblClick "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_DblClick"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_DblClick:
End Sub

Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Oops
Dim ANum
'end=35
'home=36
'insert=45
'delete=46
If Shift > 0 Then GoTo Exit_Grid1_KeyDown
If KeyCode = 35 Then 'end
    Screen.MousePointer = 11
    Grid1.Redraw = False
    NewRowIn = Grid1.Rows - 1
    NewColIn = Grid1.Col
    ChangeSel
    KeyCode = 0
    Screen.MousePointer = 0
    Grid1.Redraw = True
End If
If KeyCode = 36 Then 'home
    Screen.MousePointer = 11
    Grid1.Redraw = False
    NewRowIn = 1
    NewColIn = Grid1.Col
    ChangeSel
    KeyCode = 0
    Screen.MousePointer = 0
    Grid1.Redraw = True
End If
Grid1.Redraw = True
Screen.MousePointer = 0
GoTo Exit_Grid1_KeyDown
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_KeyDown "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_KeyDown"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_KeyDown:
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
On Error GoTo Oops
gLen = Len(Grid1.Text)
If KeyAscii = 27 Then
    Grid1.Text = Grid1.Tag
    Grid1.FixedCols = 1
    Exit Sub
End If
ColIn = Grid1.Col
RowIn = Grid1.Row
GSR = Grid1.Row
GSC = Grid1.Col
GER = Grid1.RowSel
GEC = Grid1.ColSel
'
Select Case KeyAscii
    Case 13
        If GSR <> GER Or GSC <> GEC Then
            For r = GSR To GER
                For C = GSC To GEC
                    gRow = r
                    gCol = C
                    Grid1.TextMatrix(gRow, gCol) = Scratch$
                Next C
            Next r
            Dirty = 1
        Else
            If Grid1.TextMatrix(GSR, GSC) <> Scratch$ Then
                Grid1.TextMatrix(GSR, GSC) = Scratch$
                Dirty = 1
            End If
            NewRowIn = RowIn + 1
            NewColIn = ColIn ' + 1
            If NewColIn > 200 Then
                NewColIn = 2
                NewRowIn = NewRowIn + 1
                If NewRowIn > Grid1.Rows Then
                    NewRowIn = Grid1.Rows
                End If
            End If
        End If
            IgnoreClick = True
            IgnoreTextChange = True
            Scratch$ = Grid1.TextMatrix(GSR, 1)
            InitGrid = 0
            Grid1.Row = Rowp
            Grid1.Col = 1
            Grid1.CellBackColor = vbBlue
            Grid1.CellForeColor = vbYellow
            InitGrid = 1
            Grid1.Row = NewRowIn
            Grid1.Col = 1
            Grid1.Refresh
            Exit Sub
        Case 8 'backspace
            gText = Grid1.TextMatrix(GSR, GSC)
            gLen = Len(gText)
            If gLen > 0 Then Scratch$ = Left$(gText, (gLen - 1))
            If gLen = 0 Then Scratch$ = ""
            Grid1.TextMatrix(GSR, GSC) = Scratch$
            Dirty = 1
            Exit Sub
        Case 1 'CTRL a
            If Grid1.FixedCols = 1 Then
                Grid1.FixedCols = 0
            Else
                Grid1.FixedCols = 1
            End If
            'select all columns, including col 0
            Exit Sub
        Case 3 'CTRL C
            mcopy_Click
            Exit Sub
        Case 22 'CTRL V
            mpaste_Click
            Exit Sub
        Case 24 'CTRL X
            mcut_Click
            Exit Sub
        Case 26 'CTRL Z
            Grid1.Clip = OldClip$
            Exit Sub
        Case 27 'ESC KEY
            Scratch$ = ""
            Beep
        Case Else
            Dirty = 1
            Scratch$ = Scratch$ + Chr$(KeyAscii)
            Grid1.TextMatrix(GSR, GSC) = Scratch$
            If GSC = 1 Then Grid1.TextMatrix(GSR, 1) = Scratch$
    End Select
GoTo Exit_Grid1_KeyPress
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_KeyPress "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_KeyPress"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_KeyPress:
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
txtEdit.Visible = False
Rowp = Grid1.Row
Colp = Grid1.Col
If Button = vbRightButton Then
    PopupMenu medit
Else
    'Set_TextBox
End If
End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error GoTo Oops
Dim CellData As String
Dim mHeader As String
If Button > 0 Then
    Clicking = 1
End If
If InitGrid = 0 Then Exit Sub
If Clicking = 1 Then Exit Sub
'
If txtEdit.Visible = True Then
    txtEdit.SetFocus
    Exit Sub
End If
GridX = X
GridY = y
CellX = Grid1.MouseCol
CellY = Grid1.MouseRow
r = CellY
C = CellX
If C > 0 And r > 0 Then
    CellData = Grid1.TextMatrix(r, C)
    mHeader = "Tag # " & r & ": "
    mHeader = mHeader & Trim(Grid1.TextMatrix(0, C))
    mHeader = mHeader & " = " & CellData ' & " (" & R & "," & c & ")"
    Grid1.ToolTipText = mHeader
End If
GoTo Exit_Grid1_MouseMove
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_MouseMove "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_MouseMove"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_MouseMove:
End Sub

Private Sub Grid1_EnterCell()
On Error GoTo Oops
'Debug.Print "entercell"
Dim mNumb$
If Clicking = 1 Then Exit Sub
txtEdit.Visible = False
gRow = Grid1.Row
gCol = Grid1.Col
If InitGrid = 0 Then Exit Sub
If Grid1.Col < 7 Then
    Scratch$ = Grid1.Text 'Matrix(rowp, colp)
Else
    Scratch$ = Grid1.TextMatrix(gRow, gCol)
End If
'Grid1.CellBackColor = &HC0FFFF    'lt. yellow
'Grid1.SetFocus
Grid1.Tag = Scratch$
If gRow > 0 And gRow < Grid1.Rows Then 'this used to be in leavecell
End If
GoTo Exit_Grid1_EnterCell
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_EnterCell "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_EnterCell"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_EnterCell:
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Clicking = 0
End Sub

Private Sub Grid1_RowColChange()
On Error GoTo Oops
Dim OldInit As Byte
If InitGrid = 0 Then Exit Sub
If ReplaceAll = 1 Then Exit Sub
If Clicking = 1 Then Exit Sub
'Debug.Print "RowColChange"
Rowp = Grid1.Row
Colp = Grid1.Col
Scratch$ = Grid1.TextMatrix(Rowp, Colp)
Grid1.Redraw = False
Changing = 0
OldInit = InitGrid
InitGrid = 0
Grid1.Col = 1
If Grid1.CellBackColor <> vbYellow And Grid1.CellForeColor <> vbYellow Then
    Grid1.CellBackColor = vbCyan
End If
Grid1.Col = Colp
'Debug.Print Hex(Grid1.CellBackColor)
'Debug.Print "setting  vbCyan "; Grid1.Row & ", " & Grid1.Col
InitGrid = OldInit
Grid1.Redraw = True
'Debug.Print "setting  vbCyan "; Grid1.Row & ", " & Grid1.Col
'Debug.Print "----------------------------------------------"
GoTo Exit_Grid1_RowColChange
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_RowColChange "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_RowColChange"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_RowColChange:
End Sub

Private Sub Grid1_LeaveCell()
On Error GoTo Oops
Dim OldInit As Byte
'Grid1.ToolTipText = ""
If InitGrid = 0 Then Exit Sub
If ReplaceAll = 1 Then Exit Sub
If Clicking = 1 Then Exit Sub
'Debug.Print "LeaveCell"
Dim oldx As POINTAPI
lRet = GetCursorPos(oldx)
'Debug.Print oldx.X, oldx.Y
Grid1.Redraw = False
OldInit = InitGrid
InitGrid = 0
'must be row; col; rowsel; colsel for changing cell
Grid1.Row = Rowp
Grid1.Col = 1
If Grid1.CellBackColor = vbCyan Then Grid1.CellBackColor = vbWhite
If Grid1.CellForeColor = vbYellow Then Grid1.CellBackColor = vbBlue
'Debug.Print "setting  white "; Grid1.Row & ", " & Grid1.Col
Grid1.Row = Rowp
Grid1.Col = Colp
InitGrid = OldInit
'Debug.Print "setting  white "; Grid1.Row & ", " & Grid1.Col
'Debug.Print "----------------------------------------------"
Grid1.Redraw = True
GoTo Exit_Grid1_LeaveCell
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Grid1_LeaveCell "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Grid1_LeaveCell"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Grid1_LeaveCell:
End Sub

Private Sub Grid1_SelChange()
If InitGrid = 0 Then Exit Sub
If Clicking = 1 Then Exit Sub
'Debug.Print "SelChange"
Changing = 0
Rowp = Grid1.Row
Colp = Grid1.Col
Scratch$ = Grid1.TextMatrix(Rowp, Colp)
Grid1.SetFocus
If Grid1.Row = 1 And Grid1.Col = 1 And Grid1.RowSel = Grid1.Rows - 1 And Grid1.ColSel = Grid1.Cols - 1 Then
    InitGrid = 0
    Grid1.Row = 0
    Grid1.Col = 0
    Grid1.RowSel = Grid1.Rows - 1
    Grid1.ColSel = Grid1.Cols - 1
    Grid1.Refresh
    InitGrid = 1
End If
'Debug.Print Grid1.Row, Grid1.RowSel
End Sub

Private Sub Grid1_Scroll()
txtEdit.Visible = False
End Sub

Private Sub mCopyTree_Click()
Dim aryLines() As String
Dim TreeData$
ReDim aryLines(1 To tvHeader.Nodes.Count)
For i = 1 To tvHeader.Nodes.Count
    'Debug.Print tvHeader.Nodes(i).Text
    aryLines(i) = tvHeader.Nodes(i).Text
Next i
TreeData$ = Join(aryLines, vbCrLf)
Clipboard.Clear
Clipboard.SetText TreeData$
frmMain.txtInfo.SelText = TreeData$ & vbCrLf
End Sub

Private Sub txtEdit_Change()
'Grid1.Text = Replace(txtEdit.Text, Chr$(3), "", 1, , vbTextCompare)
End Sub

Private Sub txtEdit_LostFocus()
'Grid1.Text = Replace(txtEdit.Text, Chr$(3), "", 1, , vbTextCompare)
End Sub

Private Sub txtedit_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape '27 ESC - OOPS, restore old text
        txtEdit = Grid1.Tag
        txtEdit.SelStart = Len(txtEdit)
        txtEdit.Visible = False
    Case vbKeyLeft '37 Left Arrow
        If txtEdit.SelStart = 0 And txtEdit.SelLength = 0 And Grid1.Col > 1 Then
            'Grid1.Col = Grid1.Col - 1
        Else
            If txtEdit.SelStart = 0 And txtEdit.SelLength = 0 And Grid1.Row > 1 Then
                'Grid1.Row = Grid1.Row - 1
                'Grid1.Col = Grid1.Cols - 1
            End If
        End If
    Case vbKeyUp '38 Up Arrow
        If Grid1.Row > 1 Then
            Grid1.Row = Grid1.Row - 1
        End If
    Case vbKeyRight '39 Rt Arrow
        If txtEdit.SelStart = Len(txtEdit) And Grid1.Col < Grid1.Cols - 1 Then
            'Grid1.Col = Grid1.Col + 1
        Else
            If txtEdit.SelStart = Len(txtEdit) And Grid1.Row < Grid1.Rows - 1 Then
                'Grid1.Row = Grid1.Row + 1
                'Grid1.Col = 1
            End If
        End If
    Case vbKeyDown '40 Dn Arrow
        If Grid1.Row < Grid1.Rows - 1 Then
            Grid1.Row = Grid1.Row + 1
        End If
    Case vbKeyTab
End Select
End Sub

Private Sub txtedit_KeyPress(KeyAscii As Integer)
Dim pos%, l$, r$
Select Case KeyAscii
    Case 13
        KeyAscii = 0
        Grid1.TextMatrix(Rowp, Colp) = Replace(txtEdit.Text, Chr$(3), "", 1, , vbTextCompare)
        Scratch$ = Grid1.TextMatrix(Rowp, Colp)
        txtEdit.Visible = False
        InitGrid = 0
        Grid1.Row = Rowp
        Grid1.Col = 1
        Grid1.CellBackColor = vbBlue
        Grid1.CellForeColor = vbYellow
        InitGrid = 1
        Grid1.Refresh
        ChangeSel
        '
        Rowp = Rowp + 1
        'Grid1.CellBackColor = vbWHITE
        If Rowp < Grid1.Rows Then Grid1.Row = Rowp
        Dirty = 1
        Grid1.SetFocus
    Case 8                      'BkSpc - split string @ cursor
    Case 27, 37 To 40
        Grid1 = txtEdit        'or it's going to look funny
        txtEdit.Visible = False
    Case Else
        Grid1 = txtEdit + Chr(KeyAscii)
End Select
End Sub

Private Sub Set_TextBox()   'put textbox over cell
On Error GoTo Oops
With txtEdit
    Me.Font.Name = Me.Grid1.Font.Name
    Me.Font.Size = Me.Grid1.Font.Size
    .Top = (Grid1.Top + Grid1.CellTop) - 25
    .Left = Grid1.Left + Grid1.CellLeft
    .Width = Me.TextWidth(.Text + "  ")
    If .Width < Grid1.CellWidth Then
        .Width = Grid1.CellWidth
    End If
    .Height = Grid1.CellHeight
    .Refresh
    .Text = Replace(Grid1.Text, Chr$(3), "", 1, , vbTextCompare)
    .Visible = True
    .SelStart = 0
    .SelLength = Len(.Text)
    .ZOrder 0
    '.SelStart = Len(.Text)
    .SetFocus
End With
Refresh
DoEvents
GoTo Exit_Set_TextBox
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Set_TextBox "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Set_TextBox"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Set_TextBox:
End Sub

Private Sub mCase_Click()
On Error GoTo Oops
Dim newStr$
Dim Temp As Integer
Dim Descr$
Dim dLen As Integer
Screen.MousePointer = 11
Grid1.Redraw = False
Grid1.Col = 1
GSR = Grid1.Row
GSC = Grid1.Col
GER = Grid1.RowSel
GEC = Grid1.ColSel
If GSR > GER Then
    Temp = GSR
    GSR = GER
    GER = Temp
End If
If GSC > GEC Then
    Temp = GSC
    GSC = GEC
    GEC = Temp
End If
If GEC < Grid1.Cols - 1 Then GEC = Grid1.Cols - 1
If GSR = GER Then
    tops = GSR '1
    bottoms = GER 'Grid1.Rows - 2
Else
    tops = GSR
    bottoms = GER
End If
For gRow = tops To bottoms
    gCol = 1
    Descr$ = Grid1.TextMatrix(gRow, gCol)
    dLen = Len(Descr$)
    newStr$ = StrConv(Descr$, vbProperCase)
    Grid1.TextMatrix(gRow, gCol) = newStr$
Next gRow
GoTo Exit_mCase_Click
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine mCase_Click "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in mCase_Click"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_mCase_Click:
Dirty = 1
Grid1.Redraw = True
Screen.MousePointer = 0
End Sub

Private Sub mcopyname_Click()
Clipboard.Clear
Clipboard.SetText fName$
End Sub

Private Sub mexplore_Click()
Shell "C:\Windows\Explorer.exe /n, /e, " & File1.Path, 1
End Sub

Private Sub mNote_Click()
Shell "c:\windows\mdinote.exe -h " & fName$, vbNormalFocus
End Sub

Private Sub cmdSelect_Click()
On Error GoTo Oops
fName$ = txtFile.Text
With CommonDialog1
    .CancelError = True
    .DialogTitle = "Open"
    .DefaultExt = ".wma"
    .FileName = fName$
    .Filter = "*.wma"
    .FilterIndex = 0
    .InitDir = "c:\my music"
    .ShowOpen
    fName$ = .FileName
    txtFile.Text = fName$
End With
LoadTags
GoTo Exit_Command1_Click
Oops:
If Err.Number = 32755 Then Exit Sub
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Command1_Click "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Command1_Click"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Command1_Click:
End Sub

Private Sub mPlay_Click()
fName$ = File1.Path + "\" + File1.FileName
lRet = ShellExecute(0, "play", fName, vbNullString, vbNullString, SW_SHOWNORMAL)
Caption = "Shellexecute = " & lRet
End Sub

Private Sub mSaveAs_Click()
On Error GoTo Oops
fName$ = txtFile.Text
With CommonDialog1
    .CancelError = True
    .DialogTitle = "Save as"
    .DefaultExt = "wma"
    .FileName = fName$
    .Filter = "*.wma"
    .FilterIndex = 0
    .InitDir = "c:\my music"
    .Flags = cdlOFNFileMustExist + cdlOFNOverwritePrompt
    .ShowSave
    fName$ = .FileName
End With
Debug.Print fName$
ReadGrid
WMATag.Save_Header_Tags fName$
DoEvents
Refresh
txtFile.Text = fName$
LoadTags
GoTo Exit_mSaveAs_Click
Oops:
If Err.Number = 32755 Then Exit Sub
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine mSaveAs_Click "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in mSaveAs_Click"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_mSaveAs_Click:
End Sub

Private Sub txtFile_Change()
On Error GoTo Oops
SaveSetting App.EXEName, "Settings", "LastFile", txtFile.Text
'Dir1.Path = Left$(txtFile.Text, InStrRev(txtFile.Text, "\"))
'File1.Path = Dir1.Path
'Debug.Print Dir1.Path; "\"; File1.Pattern
GoTo Exit_txtFile_Change
Oops:
If Err.Number = 76 Then Resume Next
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine txtFile_Change "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in txtFile_Change"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_txtFile_Change:
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SaveSetting App.EXEName, "Settings", "LastFile", txtFile.Text
    Dir1.Path = Left$(txtFile.Text, InStrRev(txtFile.Text, "\"))
    File1.Path = Dir1.Path
    File1.Refresh
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.List(Dir1.ListIndex)
File1.Refresh
txtFile.Text = Dir1.Path & "\" & File1.List(File1.ListIndex)
Caption = File1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Dir1.Path = Dir1.List(Dir1.ListIndex)
Dir1.SetFocus
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu mFile
Else
    Dir1.Path = Dir1.List(Dir1.ListIndex)
    txtFile.Text = File1.Path & "\" & File1.List(File1.ListIndex)
    Dir1.SetFocus
    File1.Path = Dir1.Path
    File1.Refresh
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'shoot...this executes with either mouse button down
If Ignore = 1 Then
    Ignore = 0
    Exit Sub
End If
fName$ = File1.Path + "\" + File1.FileName
txtFile.Text = fName$
LoadTags
End Sub

Private Sub File1_DblClick()
'Clipboard.Clear
'Clipboard.SetText Caption
mPlay_Click
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Oops
If KeyCode = 46 Then 'delete
    'Kill fpath
    'File1.Refresh
End If
If KeyCode = vbKeyF2 Then
    Dim OldName$
    Dim NewName$
    Dim NewPath$
    Dim extLoc As Integer
    Dim Ext$
    OldName$ = fName$
    'get extension
    extLoc = InStr(OldName$, ".")
    Ext$ = Mid$(OldName$, extLoc + 1)
    'get new name
    NewName$ = InputBox("Please enter the new File Name...", "Rename Picture", OldName$)
    If NewName$ = "" Then Exit Sub
    '
    extLoc = InStr(NewName$, ".")
    If extLoc = 0 Then
        NewName$ = NewName$ & "." & Ext$
    End If
    NewPath$ = File1.Path & "\" & NewName$
    NewPath$ = Replace(NewPath$, "\\", "\")
    fName$ = File1.Path & "\" & fName$
    Name fName$ As NewPath$
    File1.Refresh
    'Clipboard.Clear
    'Clipboard.SetText NewName$
End If
GoTo Exit_File1_KeyDown
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine File1_KeyDown "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in File1_KeyDown"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_File1_KeyDown:
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    File1_Click
End If
'If KeyAscii = 27 Then fName$ = "": WaveName$ = ""
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
lXPoint = CLng(X / Screen.TwipsPerPixelX)
lYPoint = CLng(y / Screen.TwipsPerPixelY)
Ignore = 1
If Button = 2 Then
    lIndex = SendMessage(File1.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
    File1.ListIndex = lIndex
    fName$ = File1.Path + "\" + File1.List(File1.ListIndex)
    PopupMenu mFile
    Exit Sub
Else
    fName$ = File1.Path + "\" + File1.FileName
    txtFile.Text = fName$
    LoadTags
    'Clipboard.Clear
    'Clipboard.SetText WaveName$
End If
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
' present related tip message
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
With File1
    lXPoint = CLng(X / Screen.TwipsPerPixelX)
    lYPoint = CLng(y / Screen.TwipsPerPixelY)
    If Button = 0 Then ' if no button was pressed
        ' get selected item from list
        lIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
        ' show tip or clear last one
        If (lIndex >= 0) And (lIndex <= .ListCount) Then
            .ToolTipText = .List(lIndex) & " - Double click to copy to Clipboard..."
        Else
            .ToolTipText = ""
        End If
    End If '(button=0)
End With '(List1)
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'Clipboard.Clear
'Clipboard.SetText Caption
End Sub

Sub LoadTags()
On Error GoTo Oops
DoEvents
Set WMATag = New TagReader.InfoTag
ClearGrid
TagSuccess = WMATag.Find_Header_Tags(fName$)
Caption = "Retrieving Tag Data for " & fName$ & "..."
If TagSuccess = True Then
    With Grid1
        .Redraw = False
        .Visible = False
        .TextMatrix(1, 1) = WMATag.Title
        .TextMatrix(2, 1) = WMATag.Artist
        .TextMatrix(3, 1) = WMATag.Album
        .TextMatrix(4, 1) = WMATag.AlbumArtist
        .TextMatrix(5, 1) = WMATag.Composer
        .TextMatrix(6, 1) = WMATag.Publisher
        .TextMatrix(7, 1) = WMATag.FileSize
        .TextMatrix(8, 1) = WMATag.TrackNumber '
        .TextMatrix(9, 1) = WMATag.Year
        .TextMatrix(10, 1) = Format(WMATag.Bitrate, "###,#")    '
        .TextMatrix(11, 1) = Format(WMATag.Frequency, "###,#")
        .TextMatrix(12, 1) = WMATag.IsCopyright    '
        .TextMatrix(13, 1) = WMATag.Copyright
        .TextMatrix(14, 1) = WMATag.IsLicensed    '
        .TextMatrix(15, 1) = WMATag.IsProtected
        .TextMatrix(16, 1) = WMATag.Rating
        .TextMatrix(17, 1) = WMATag.Duration
        .TextMatrix(18, 1) = WMATag.Mode
        .TextMatrix(19, 1) = WMATag.Comments
        'byte #----------------------------------------------------------------------------------------------->
        ' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30
        '------------|-------------|-------------|-------------|------------|-------------|-------------|------
        '30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02
        '            |             |             |             |^^          |             |             |
        '31-------34 | 35-------38 | 39-------42 | 43-------46 |47-------50 | 51-------54 | 55-------58 | 59 60
        '33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00
        '                                                       ^^                          ^^    ^^
        '                                                                               TitleLen ArtistLen
        '_______________________________________________________________________________________________________
        'byte 65 is the song title with a len specified by byte 55.
        'immediately after that is the artist's name.  After the artist's name is the following
'        Dim ch$
'        Dim chNum As Integer
'        Dim i As Integer
'        Dim HexStr$
'        .SelText = "Header Hex  = " & vbTab
'        For i = 1 To 60
'            HexStr$ = Right$("0" & (Hex$(Asc(Mid$(WMATag.FullHeader, i, 1)))), 2)
'            .SelText = HexStr$ & " "
'            If i = 30 Then
'                .SelText = vbCrLf
'                .SelText = "Header Hex  = " & vbTab
'            End If
'        Next i
        Debug.Print "Setting Tag Info..."
        Debug.Print WMATag.Info; vbCrLf
        txtInfo.SelText = WMATag.Info & vbCrLf
        '        .SelText = "Header(" & Len(WMATag.FullHeader) & ") = " & vbTab
        '        .SelText = Replace(WMATag.FullHeader, Chr$(0), ".") & vbCrLf
        .Redraw = True
        .Visible = True
        DoEvents
        Refresh
    End With
Else
    Grid1.Text = "WMA Tag Read Error"
    Debug.Print "WMA Tag Read Error"
End If
GoTo Exit_LoadTags
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine LoadTags "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in LoadTags"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_LoadTags:
Caption = "Tag Data complete for " & fName$ & "."
End Sub

Sub ClearGrid()
On Error GoTo Oops
With Grid1
    .Redraw = False
    .Visible = False
    .FillStyle = flexFillRepeat
    .Row = 1
    .Col = 1
    .RowSel = 19
    .ColSel = 1
    .Text = ""
    .CellBackColor = vbWhite
    .CellForeColor = vbBlack
    .FillStyle = flexFillSingle
    .RowSel = 1
    .Redraw = True
    .Visible = True
    DoEvents
    Refresh
End With
GoTo Exit_ClearGrid
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ClearGrid "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ClearGrid"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ClearGrid:
End Sub

Sub ChangeSel()
On Error GoTo Oops
'scratch$ = ""
If NewRowIn = 0 And NewColIn = 0 Then Exit Sub
If NewRowIn < 0 Then Exit Sub
If NewColIn < 0 Then Exit Sub
Grid1.Row = NewRowIn
Grid1.Col = NewColIn
Rowp = Grid1.Row
Colp = Grid1.Col
Grid1.RowSel = NewRowIn
Grid1.ColSel = NewColIn
ColIn = NewColIn
RowIn = NewRowIn
If Grid1.RowIsVisible(NewRowIn) = False Then
    Grid1.TopRow = NewRowIn
End If
Refresh
GoTo Exit_ChangeSel
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ChangeSel "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ChangeSel"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ChangeSel:
End Sub

Private Sub mcopy_Click()
Clip$ = Grid1.Text
gClip$ = ""
'
ColIn = Grid1.Col
RowIn = Grid1.Row
GSR = Grid1.Row
GSC = Grid1.Col
GER = Grid1.RowSel
GEC = Grid1.ColSel
'
CopySR = GSR
CopyER = GER
CopySC = GSC
CopyEC = GEC
gClip$ = Grid1.Clip
Clipboard.Clear
Clipboard.SetText gClip$
Debug.Print Str$(GER - GSR + 1); " Rows copied"
End Sub

Private Sub mcut_Click()
Screen.MousePointer = 11
Grid1.Redraw = False
Clip$ = Grid1.Text
ColIn = Grid1.Col
RowIn = Grid1.Row
GSR = Grid1.Row
GSC = Grid1.Col
GER = Grid1.RowSel
GEC = Grid1.ColSel
'
CopySR = GSR
CopyER = GER
CopySC = GSC
CopyEC = GEC
gClip$ = Grid1.Clip
ClearSels Grid1
Clipboard.Clear
Clipboard.SetText gClip$
Screen.MousePointer = 0
Grid1.Redraw = True
End Sub

Private Sub mpaste_Click()
On Error GoTo Oops
Dim rFound As Integer
Dim cFound As Integer
Dim OneCell As Byte
Dim PasteRange As Byte
Dim CLen As Long
Screen.MousePointer = 11
Grid1.Redraw = False
InitGrid = 0
ColIn = Grid1.Col
RowIn = Grid1.Row
GSR = Grid1.Row
GSC = Grid1.Col
GER = Grid1.RowSel
GEC = Grid1.ColSel
OldClip$ = ""
'strip off lf from clipboard
ClipIn$ = Clipboard.GetText
NewClip = PastePre
CLen = Len(ClipIn$)
If CLen < 5000 Then GoTo StringClip
'********************************************************************************
ByteClip:
Dim ClipChar As Long
Dim CharByte() As Byte
Dim ClipByte() As Byte
Dim ClipPre() As Byte
Dim ClipSuf() As Byte
Dim ClipByteLen As Long
Dim ClipBytePos As Long
Dim AddCount As Integer
Dim SufLen As Integer
Dim PreLen As Integer
CharByte() = StrConv(ClipIn$, vbFromUnicode)
ClipPre() = StrConv(PastePre, vbFromUnicode)
ClipSuf() = StrConv(PasteSuf, vbFromUnicode)
SufLen = UBound(ClipSuf)
PreLen = UBound(ClipPre)
If SufLen < 0 Then SufLen = 0
If PreLen < 0 Then PreLen = 0
ClipByteLen = 0
ClipBytePos = 0 - 1
For ClipChar = 0 To CLen - 1
    If chN = 13 And ClipChar < CLen Then
        If PreLen > 0 Then
            'NewClip = NewClip & PastePre
            ClipByteLen = ClipByteLen + PreLen
            ReDim Preserve ClipByte(ClipByteLen)
            For AddCount = 0 To UBound(ClipPre) - 1
                ClipBytePos = ClipBytePos + 1
                ClipByte(ClipBytePos) = ClipPre(AddCount)
            Next AddCount
        End If
    End If
    chIn$ = Chr$(CharByte(ClipChar))
    chN = Asc(chIn$)
    If chN = 13 Then
        If SufLen > 0 Then
            'NewClip = NewClip & PasteSuf & chin$
            ClipByteLen = ClipByteLen + SufLen
            ReDim Preserve ClipByte(ClipByteLen)
            For AddCount = 0 To UBound(ClipSuf) - 1
                ClipBytePos = ClipBytePos + 1
                ClipByte(ClipBytePos) = ClipSuf(AddCount)
            Next AddCount
        End If
        ClipBytePos = ClipBytePos + 1
        ReDim Preserve ClipByte(ClipBytePos)
        ClipByte(ClipBytePos) = CharByte(ClipChar)
        GoTo cr1
    End If
    If chN <> 10 Then
        'NewClip = NewClip & chin$
        ClipBytePos = ClipBytePos + 1
        ReDim Preserve ClipByte(ClipBytePos)
        ClipByte(ClipBytePos) = CharByte(ClipChar)
   End If
cr1:
    pVal = ((ClipChar * 100) \ CLen)
    If pVal > LastpVal Then
        Caption = "Parsing Clipboard Character " & Str$(ClipChar) & " of " & CLen
        DoEvents
        LastpVal = pVal
    End If
Next ClipChar
If UBound(ClipSuf) > 0 Then
    'NewClip = NewClip & PasteSuf
    ClipByteLen = ClipByteLen + UBound(ClipSuf)
    ReDim Preserve ClipByte(ClipByteLen)
    For AddCount = 0 To UBound(ClipSuf) - 1
        ClipBytePos = ClipBytePos + 1
        ClipByte(ClipBytePos) = ClipSuf(AddCount)
    Next AddCount
End If
'
'added 1 to clipbytepos because it was missing the last character
NewClip = Space(ClipBytePos + 1)
'should use AllocStrB instead of space
CopyMemory ByVal NewClip, ByVal VarPtr(ClipByte(0)), ClipBytePos + 1
'
GoTo BuildClip
'********************************************************************************
StringClip:
For ClipChar = 1 To CLen
    If chN = 13 And ClipChar < CLen Then
        NewClip = NewClip & PastePre
    End If
    chIn$ = Mid$(ClipIn$, ClipChar, 1)
    chN = Asc(chIn$)
    If chN = 13 Then
        NewClip = NewClip & PasteSuf & chIn$ '& PastePre
        GoTo cr
    End If
    If chN <> 10 Then
        NewClip = NewClip & chIn$
   End If
cr:
    pVal = ((ClipChar * 100) \ CLen)
    If pVal > LastpVal Then
        Caption = "Parsing Clipboard Character " & Str$(ClipChar) & " of " & CLen
        DoEvents
        LastpVal = pVal
    End If
Next ClipChar
NewClip = NewClip & PasteSuf
'********************************************************************************
BuildClip:
'MsgBox Timer - sTime#
'
Dim Extry As Integer
If Len(NewClip) > 0 Then Extry = Asc(Right$(NewClip, 1))
If Extry = 13 Then
    ClipIn$ = Left(NewClip, (Len(NewClip) - 1))
Else
    ClipIn$ = NewClip
End If
'
'check rows and columns in clipboard
Dim aryRows
Dim aryCols
'changed below from 0 to -1 because 0 means the array is valid
rFound = -1
cFound = -1
aryRows = Split(ClipIn$, vbCr)
If IsEmpty(aryRows) = False Then
    If UBound(aryRows) > 0 Then rFound = UBound(aryRows)
End If
If rFound >= 0 Then aryCols = Split(aryRows(0), vbTab)
If IsEmpty(aryCols) = False Then
    If UBound(aryCols) > 0 Then cFound = UBound(aryCols)
End If
'
'below added 10/2003 because excel had an ubound for arycols of -1
If rFound < 0 Then rFound = 0
If cFound < 0 Then cFound = 0
'when 1 row is selected, it puts in a CR!
'
OneCell = 0
PasteRange = 0
If CopyER = 0 And CopySR = 0 And CopyEC = 0 And CopySC = 0 Then OneCell = 1
If GER > GSR Or GEC > GSC Then PasteRange = 1
'copy cell to range if paste range>copy range
If OneCell = 1 And PasteRange = 1 Then
    'copy cell to range if paste range>copy range
'If (CopyER - CopySR = 0) And (CopyEC - CopySC = 0) And (GSR <= GER Or GSC <= GEC) Then
    OldClip$ = Grid1.Clip
    Grid1.FillStyle = flexFillRepeat
    Grid1.Row = GSR
    Grid1.Col = GSC
    Grid1.RowSel = GER
    Grid1.ColSel = GEC
    Grid1.Text = gClip$
    Grid1.FillStyle = flexFillSingle
ElseIf CopyER - CopySR = 0 And GSR < GER And GSC <= GEC Then
    'copy row to range of rows
    OldClip$ = ""
    For r = GSR To GER
        Grid1.Row = r
        Grid1.Col = GSC
        Grid1.RowSel = r
        Grid1.ColSel = GSC + (CopyEC - CopySC)
        OldClip$ = OldClip$ + vbCrLf + Grid1.Clip
        Grid1.Clip = gClip$
    Next r
ElseIf GER >= GSR And GEC >= GSC Then
    Grid1.Row = GSR
    Grid1.Col = GSC
    Grid1.RowSel = GSR + (CopyER - CopySR)
    Grid1.ColSel = GSC + (CopyEC - CopySC)
    OldClip$ = Grid1.Clip
    'clipin$ = Clipboard.GetText
    'If Grid1.RowSel > GSR Then GoTo paste
    'If Grid1.ColSel > GSC Then GoTo paste
    'Debug.Print clipin
    If rFound > 0 Or cFound > 0 Then
        Grid1.Row = GSR
        Grid1.Col = GSC
        Grid1.RowSel = GSR + rFound
        Grid1.ColSel = GSC + cFound
        OldClip$ = Grid1.Clip
    End If
paste:
    Grid1.Clip = ClipIn$
Else
    Grid1.Col = GSC
    Grid1.ColSel = GSC + (CopyEC - CopySC)
    OldClip$ = Grid1.Clip
    Grid1.Clip = Clipboard.GetText 'Gclip$
End If
'
Dim PR As Integer
InitGrid = 0
For PR = GSR To GSR + rFound + 1
    Grid1.Row = PR
    Grid1.Col = 1
    Grid1.CellBackColor = vbBlue
    Grid1.CellForeColor = vbYellow
Next PR
InitGrid = 1
Grid1.Refresh
ChangeSel
GoTo Exit_mpaste_Click
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine mpaste_Click "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in mpaste_Click"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_mpaste_Click:
'tried to get the message to display after paste
If GSR + rFound < Grid1.Rows Then Grid1.RowSel = GSR + rFound
If GSC + cFound < Grid1.Cols Then Grid1.ColSel = GSC + cFound
Scratch$ = Grid1.TextMatrix(GSR, GSC)
Grid1.Redraw = True
Grid1.Visible = True
Grid1.Refresh
DoEvents
InitGrid = 1
Dim endr%
If rFound > 0 Then
    endr% = GSR + rFound
    If endr% > Grid1.Rows - 2 Then endr% = Grid1.Rows - 2
Else
    endr% = GER
End If
Grid1.Col = GSC
Dirty = 1
Screen.MousePointer = 0
End Sub

Sub ReadGrid()
With Grid1
    .Redraw = False
    WMATag.Title = .TextMatrix(1, 1)
    WMATag.Artist = .TextMatrix(2, 1)
    WMATag.Album = .TextMatrix(3, 1)
    WMATag.AlbumArtist = .TextMatrix(4, 1)
    WMATag.Composer = .TextMatrix(5, 1)
    WMATag.Publisher = .TextMatrix(6, 1)
    WMATag.TrackNumber = .TextMatrix(8, 1)
    WMATag.Year = .TextMatrix(9, 1)
    WMATag.IsCopyright = False  '=.TextMatrix(12, 1)
    WMATag.Copyright = .TextMatrix(13, 1)
    WMATag.IsLicensed = False     '=.TextMatrix(14, 1)
    WMATag.IsProtected = False  '= .TextMatrix(15, 1)
    WMATag.Rating = .TextMatrix(16, 1)
    WMATag.Duration = .TextMatrix(17, 1)
    WMATag.Comments = .TextMatrix(19, 1)
    .Redraw = True
    .Visible = True
    DoEvents
    Refresh
End With
End Sub
