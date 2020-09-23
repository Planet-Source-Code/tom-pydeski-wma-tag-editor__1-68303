Attribute VB_Name = "FlexGrid"
Option Explicit
Public Const EM_UNDO = &HC7
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302
Global GridBeg As Integer
Global RowIn As Integer
Global ColIn As Integer
Global NewRowIn As Integer
Global NewColIn As Integer
Global Rowp As Integer
Global Colp As Integer
Global Clip$
Global ArrayMax As Integer
Global RecNo As Long
Global MaxRecNo As Long
Global Clips$
Global GSR As Integer
Global GER As Integer
Global GSC As Integer
Global GEC  As Integer
Global CopySR As Integer
Global CopyER As Integer
Global CopySC As Integer
Global CopyEC  As Integer
Global gClip$ '5/97
Global OldClip$
Global ClipIn$
Global Clicking As Byte
Global Dirty As Byte
'sdirty is schedule grid for stats program
Global sDirty As Byte
Global VisibleGridRows As Byte
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
'Public Declare Function SetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'below used for grid find
Global FoundPos As Long
Global gFindString$
Global gRepString$
Global FoundLine As Integer
Global LastLine As Integer
Global NewFound As Integer
Global gCurPos As Long
Global Finding As Byte
Global TotalFounds As Integer
Global Source$
Global FoundText
Global Found As Integer
Global tops As Integer
Global bottoms As Integer
Global ReplaceAll As Byte
Global DoingFind As Byte
'**************************************
'Windows API/Global Declarations for :Auto Scroll ListBox/MSFlexGrid
'**************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Constant
Global Const WM_VSCROLL = &H115
Global Const SB_BOTTOM = 7
'Call below where auto scroll is intended
'SendMessage grid1.hwnd, WM_VSCROLL, SB_BOTTOM, 0
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Dim gText$
Dim i As Integer
'=============================
Global Scratch$
Global r As Integer
Global C As Integer
Global gLen As Integer
Global gRow As Integer
Global gCol As Integer
'made above global for wma tag editor 3/2007
'=============================
Dim NewClip$
Dim RowClip$
Dim TempClip$
'below put in here for use in slc tcp
Global InitGrid As Byte
Global GridX As Integer
Global GridY As Integer
Global CellX As Integer
Global CellY As Integer
Global Editing As Byte
Global PastePre As String, PasteSuf As String

Sub ClearSels(Grid As MSFlexGrid)
OldClip$ = Grid.Clip
Grid.FillStyle = flexFillRepeat
'
'below will clear ALL Sels
'Grid.Row = 1
'Grid.Col = 1
'Grid.RowSel = Grid.Rows - 1
'Grid.ColSel = Grid.cols - 1
Grid.Text = ""
Grid.FillStyle = flexFillSingle
'rowsize = Abs(Grid.RowSel - Grid.Row)
'colsize = Abs(Grid.ColSel - Grid.Col)
'NewClip$ = ""
'For r = 0 To rowsize
'    If colsize > 0 Then
'        NewClip$ = NewClip$ + String(colsize, 9)
'    End If
'    If rowsize > 0 Then NewClip$ = NewClip$ + vbCr
'Next r
'Grid.Clip = NewClip$
Scratch$ = ""
End Sub

Sub ClearALLSels(Grid As MSFlexGrid)
OldClip$ = Grid.Clip
Grid.FillStyle = flexFillRepeat
'below will clear ALL Sels
Grid.Row = 1
Grid.Col = 1
Grid.RowSel = Grid.Rows - 1
Grid.ColSel = Grid.Cols - 1
Grid.Text = ""
Grid.FillStyle = flexFillSingle
Scratch$ = ""
End Sub

Sub ChangeSel(Grid As MSFlexGrid)
On Error GoTo Oops
Dim topMax As Integer
Scratch$ = ""
If NewRowIn >= Grid.Rows Then Exit Sub
If NewColIn >= Grid.Cols Then Exit Sub
If NewRowIn = 0 Or NewColIn = 0 Then Exit Sub
If NewRowIn < 0 Then Exit Sub
If NewColIn < 0 Then Exit Sub
Grid.Row = NewRowIn
Grid.Col = NewColIn
Grid.RowSel = NewRowIn
Grid.ColSel = NewColIn
ColIn = NewColIn
RowIn = NewRowIn
topMax = (Grid.Rows - (Grid.Height \ Grid.RowHeight(1))) + 1
If Grid.RowIsVisible(NewRowIn) = False Then
    Grid.TopRow = NewRowIn
End If
If RowIn > topMax And topMax > 0 Then
    'grid.TopRow = topmax
Else
    'grid.TopRow = RowIN
End If
If DoingFind = 1 Then Grid.TopRow = NewRowIn
Grid.Refresh
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

Sub GridCopy(Grid As MSFlexGrid)
If Grid.Visible = False Then Exit Sub
ColIn = Grid.Col
RowIn = Grid.Row
GSR = Grid.Row
GER = Grid.RowSel
GSC = Grid.Col
GEC = Grid.ColSel
'
CopySR = GSR
CopyER = GER
CopySC = GSC
CopyEC = GEC
gClip$ = Grid.Clip
Clipboard.Clear
Clipboard.SetText gClip$
End Sub

Sub GridPaste(Grid As MSFlexGrid)
On Error GoTo Oops
Dim CLen As Long
Dim Extry As Integer
Dim rFound As Integer
Dim cFound As Integer
Dim OneCell As Byte
Dim PasteRange As Byte
If Grid.Visible = False Then Exit Sub
Screen.MousePointer = 11
Clicking = 1
Grid.Redraw = False
InitGrid = 0
ColIn = Grid.Col
RowIn = Grid.Row
GSR = Grid.Row
GSC = Grid.Col
GER = Grid.RowSel
GEC = Grid.ColSel
OldClip$ = ""
'strip off lf from clipboard
ClipIn$ = Clipboard.GetText
CLen = Len(ClipIn$)
NewClip$ = Replace(ClipIn$, vbLf, "")
'NewClip$ = NewClip$ + PasteSuf
Extry = Asc(Right$(NewClip$, 1))
If Extry = 13 Then
    ClipIn$ = Left(NewClip$, (Len(NewClip$) - 1))
Else
    ClipIn$ = NewClip$
End If
'
'check rows and columns in clipboard
rFound = 0
cFound = 0
Dim aryRows
Dim aryCols
aryRows = Split(ClipIn$, vbCr)
If IsEmpty(aryRows) = False Then
    rFound = UBound(aryRows)
End If
If rFound >= 0 Then aryCols = Split(aryRows(0), vbTab)
If IsEmpty(aryCols) = False Then
    'does this put in an extra column?
    cFound = UBound(aryCols)
End If
'
'when 1 row is selected, it puts in a CR!
'rfound = rfound - 1
'If rfound < 0 Then rfound = 0
Grid.Redraw = False
'
OneCell = 0
PasteRange = 0
If CopyER = 0 And CopySR = 0 And CopyEC = 0 And CopySC = 0 Then OneCell = 1
If GER > GSR Or GEC > GSC Then PasteRange = 1
'copy cell to range if paste range>copy range
If OneCell = 1 And PasteRange = 1 Then
'If (CopyER - CopySR = 0) And (CopyEC - CopySC = 0) And (GSR <= GER Or GSC <= GEC) Then
    OldClip$ = Grid.Clip
    Grid.FillStyle = flexFillRepeat
    Grid.Row = GSR
    Grid.Col = GSC
    Grid.RowSel = GER
    Grid.ColSel = GEC
    Grid.Text = gClip$
    Grid.FillStyle = flexFillSingle
ElseIf CopyER - CopySR = 0 And GSR < GER And GSC <= GEC Then
    'copy row to range of rows
    OldClip$ = ""
    For r = GSR To GER
        Grid.Row = r
        Grid.Col = GSC
        Grid.RowSel = r
        Grid.ColSel = GSC + (CopyEC - CopySC)
        OldClip$ = OldClip$ & vbCrLf & Grid.Clip
        Grid.Clip = gClip$
    Next r
ElseIf GER >= GSR And GEC >= GSC Then
    Grid.Row = GSR
    Grid.Col = GSC
    Grid.RowSel = GSR + (CopyER - CopySR)
    Grid.ColSel = GSC + (CopyEC - CopySC)
    OldClip$ = Grid.Clip
    'ClipIn$ = Clipboard.GetText
    'If Grid.RowSel > GSR Then GoTo paste
    'If Grid.ColSel > GSC Then GoTo paste
    'Debug.Print ClipIn$
    If rFound > 0 Or cFound > 0 Then
        Grid.Row = GSR
        Grid.Col = GSC
        If GSR + rFound > Grid.Rows Then
            Grid.Rows = Grid.Rows + GSR + rFound
            Grid.Height = ((Grid.RowHeight(0) + 30) * (Grid.Rows - 1))
        End If
        Grid.RowSel = GSR + rFound
        'the -1 was added 6/2002 because it was putting an extra column
        'but i had to take it back out 7/9/2002 because it was taking away
        'one of the columnsx
        Dim newCols As Integer
        newCols = (GSC + cFound) ' - 1
        If newCols > Grid.Cols Then
            newCols = (GSC + cFound) - 1
        End If
        Grid.ColSel = newCols
        OldClip$ = Grid.Clip
    End If
paste:
    Grid.Clip = ClipIn$
Else
    Grid.Col = GSC
    Grid.ColSel = GSC + (CopyEC - CopySC)
    OldClip$ = Grid.Clip
    Grid.Clip = Clipboard.GetText 'Gclip$
End If
'For i = 1 To Len(clipin)
'    chin$ = Mid$(clipin, i, 1)
'    chn = Asc(chin$)
'    Debug.Print i, chn, chin$
'Next i
'BEEEP
Grid.Row = GSR
Grid.Col = GSC
Grid.RowSel = GSR + rFound
'the -1 was added 6/2002 because it was putting an extra column
If GSC > 0 Then Grid.ColSel = (GSC + cFound) - 1
Scratch$ = Grid.TextMatrix(GSR, GSC)
'Debug.Print Scratch$
GoTo Exit_GridPaste
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GridPaste "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GridPaste"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GridPaste:
Grid.Redraw = True
Grid.Refresh
InitGrid = 1
Screen.MousePointer = 0
End Sub

Sub SizeGrid(Grid As MSFlexGrid, gFormat As String)
fixsize:
Dim cWidth As Integer
Dim gridw As Integer
Dim C As Integer
cWidth = 130
'The FormatString property contains segments separated by pipe characters (|).
'The text between pipes defines a column and may also contain special
'alignment characters. These characters align the entire column to the
'left (<), center (^), or right (>).
'In addition, the text is assigned to row zero by default, and the text width
'defines the width of each column.
'The FormatString property may contain a semicolon (;).
'This causes the remainder of the string to be interpreted as row heading and
'row width information. In addition, the text is assigned to column zero by default,
'and the longest string defines the width of column zero.
'
'usage sizegrid "<Game#|<Date |<Day|<Opponent          | <Time | <TV |^Result|<Score|^Record
                  '162, Apr 2,Mon., Kansas City Royals , 1:05 PM, Fox5, Won,    7-3,1-0
' Set column headers.
'format string no longer shrinks columns, so i had to add below
For C = 0 To Grid.Cols - 1
    Grid.ColWidth(C) = 100
Next C
Grid.FormatString = gFormat
GoTo news
Grid.ColWidth(0) = 6 * cWidth 'mess num
Grid.ColWidth(1) = 3 * cWidth 'print
Grid.ColWidth(2) = 6 * cWidth 'Size
Grid.ColWidth(3) = 7 * cWidth 'Text Color
Grid.ColWidth(4) = 3 * cWidth 'Text Blink
Grid.ColWidth(5) = 6 * cWidth 'Back Color
Grid.ColWidth(6) = 3 * cWidth 'Back Blink
Grid.ColWidth(7) = 78 * cWidth
news:
gridw = 0
For i = 0 To Grid.Cols - 1
    gridw = gridw + Grid.ColWidth(i)
Next i
Grid.Width = gridw + 450
'grid.Top = 700
'Grid.Height = Grid.Parent.ScaleHeight - (Grid.Top + 100)
'MessLoc.Width = Me.Width - MessLoc.Left
'MessLoc.Left = grid.Left + grid.Width
'MessLoc.Top = grid.Top + 300
'MessLoc.Height = grid.Height - 600
End Sub

Sub GridKeyDown(Grid As MSFlexGrid, KeyCode As Integer, Shift As Integer)
On Error GoTo Oops
Dim Rowout As Integer
Dim lRet As Long
'end=35
'home=36
'insert=45
'delete=46
Debug.Print "key="; KeyCode, Shift
If KeyCode = 45 Then 'insert=45
    Screen.MousePointer = 11
    Grid.Redraw = False
    GSR = Grid.Row
    Grid.Redraw = False
    InitGrid = 0
    '
inserts:
    Dim InsRows As Integer
    InsRows = InputBox("How many rows do you want to add", "Insert Rows...", 1)
    If Val(InsRows) > 0 Then
        AddRows Grid, InsRows
    End If
    '
    Screen.MousePointer = 0
    Grid.Row = GSR
    Grid.Redraw = True
    InitGrid = 1
    If LCase(Grid.Name) = "grid2" Then
        sDirty = 1
    Else
        Dirty = 1
    End If
End If
If KeyCode = 46 Then 'delete
    If LCase(Grid.Name) = "grid2" Then
        sDirty = 1
    Else
        Dirty = 1
    End If
    Screen.MousePointer = 11
    Grid.Redraw = False
    ClearSels Grid
    If Grid.Col = 1 And Grid.ColSel > 1 Then
        GSR = Grid.Row
        GSC = Grid.Col
        GER = Grid.RowSel
        GEC = Grid.ColSel
        eMess$ = "Do you want to delete these " & Str$(GER - GSR + 1) & " rows from the spreadsheet?"
        lRet = MsgBox(eMess$, vbYesNoCancel, "Delete Rows")
        If lRet = vbCancel Then Grid.Clip = OldClip$
        If lRet = vbYes Then
            Screen.MousePointer = 11
            InitGrid = 0
            Grid.Redraw = False
            For Rowout = GSR To GER
                'remove the starting row to ending row
                'if we remove rowout, we will overshoot ger
                'therfore we always remove row gsr
                Grid.RemoveItem GSR
            Next Rowout
            'fix the numbers in column 0
            'now gsr is the first undeleted row
            For Rowp = GSR To Grid.Rows - 1
                If Val(Grid.TextMatrix(Rowp, 0)) = 1 Then
                    Exit For
                End If
                Grid.TextMatrix(Rowp, 0) = (Grid.TextMatrix(Rowp - 1, 0)) + 1
            Next Rowp
            '
            Grid.Row = GSR - 1
            Grid.Col = 1
            Grid.RowSel = GSR - 1
            Grid.ColSel = 1
            Rowp = 1
            Colp = 1
            NewRowIn = GSR
            NewColIn = 1
            ChangeSel Grid
            InitGrid = 1
            Grid.Redraw = True
        End If
    End If
    Grid.Refresh
    Grid.Redraw = True
    Screen.MousePointer = 0
End If
If KeyCode = 9 Then 'Or KeyCode = 39 Then
    If Shift = 0 Then
        NewColIn = ColIn + 1
        NewRowIn = RowIn
    Else
        NewColIn = ColIn - 1
        NewRowIn = RowIn
    End If
    ChangeSel Grid
End If
If Shift = 2 And KeyCode = 67 Then 'CTRL C
    GridCopy Grid
    Exit Sub
End If
If Shift = 2 And KeyCode = 86 Then 'CTRL V
    'mpaste_Click
    Exit Sub
End If
If Shift = 2 And KeyCode = 88 Then 'CTRL X
    GridCopy Grid
    ClearSels Grid
    Exit Sub
End If
GoTo Exit_GridKeyDown
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GridKeyDown "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GridKeyDown"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GridKeyDown:
End Sub

Sub GridKeyPress(Grid As MSFlexGrid, KeyAscii As Integer)
On Error GoTo Oops
'Dirty = 1
If KeyAscii = 27 Then
    Scratch$ = ""
    Clicking = 0
    If GSR < GER And GSC < GEC Then Exit Sub
'    If OldClip$ = grid.text Then meditparm_Click
'    If OldClip$ <> "" Then grid.text = OldClip$
    If OldClip$ <> "" Then Grid.Clip = OldClip$: Exit Sub
    Grid.Visible = False
End If
'
Select Case KeyAscii
    Case 13
        If GSR <> GER Or GSC <> GEC Then
            For r = GSR To GER
                For C = GSC To GEC
                    Grid.Row = r
                    Grid.Col = C
                    Grid.Text = Scratch$
                Next C
            Next r
            'Dirty = 1
        Else
            Grid.Text = Scratch$
            If Grid.Col < Grid.Cols - 1 Then
                NewRowIn = RowIn
                NewColIn = ColIn + 1
            Else
                NewRowIn = RowIn + 1
                NewColIn = ColIn
            End If
            ChangeSel Grid
        End If
        Exit Sub
    Case 8
        gText$ = Grid.TextMatrix(GSR, GSC)
        gLen = Len(gText$)
        If gLen > 0 Then Scratch$ = Left$(gText$, (gLen - 1))
        If gLen = 0 Then Scratch$ = ""
        Grid.TextMatrix(GSR, GSC) = Scratch$
        If LCase(Grid.Name) = "grid2" Then
            sDirty = 1
        Else
            Dirty = 1
        End If
        Exit Sub
    Case 3 'CTRL C
        GridCopy Grid
        Exit Sub
    Case 22 'CTRL V
        GridPaste Grid
        Grid.Refresh
        If LCase(Grid.Name) = "grid2" Then
            sDirty = 1
        Else
            Dirty = 1
        End If
        Exit Sub
    Case 24 'CTRL X
        GridCopy Grid
        ClearSels Grid
        Grid.Refresh
        Exit Sub
    Case 26 'CTRL Z
        Grid.Clip = OldClip$
        If LCase(Grid.Name) = "grid2" Then
            sDirty = 1
        Else
            Dirty = 1
        End If
        Exit Sub
    Case 27 'ESC KEY
        Scratch$ = ""
        Beep
    Case Else
        If LCase(Grid.Name) = "grid2" Then
            sDirty = 1
        Else
            Dirty = 1
        End If
        Scratch$ = Scratch$ + Chr$(KeyAscii)
        Grid.TextMatrix(GSR, GSC) = Scratch$
        'Debug.Print KeyAscii
    End Select
GoTo Exit_GridKeyPress
Oops:
If Err = 32755 Then GoTo Exit_GridKeyPress
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GridKeyPress "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GridKeyPress"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GridKeyPress:
End Sub

Sub txtKeyDown(txtEdit As Control, KeyCode As Integer, Shift As Integer, Grid As MSFlexGrid)
Select Case KeyCode
    Case 27     'ESC - OOPS, restore old text
        txtEdit = Grid.Tag
        txtEdit.SelStart = Len(txtEdit)
    Case 37     'Left Arrow
        If txtEdit.SelStart = 0 And txtEdit.SelLength = 0 And Grid.Col > 1 Then
            'Grid.Col = Grid.Col - 1
        Else
            If txtEdit.SelStart = 0 And txtEdit.SelLength = 0 And Grid.Row > 1 Then
                'Grid.Row = Grid.Row - 1
                'Grid.Col = Grid.cols - 1
            End If
        End If
    Case 38     'Up Arrow
        If Grid.Row > 1 Then
            Grid.Row = Grid.Row - 1
        End If
    Case 39     'Rt Arrow
        If txtEdit.SelStart = Len(txtEdit) And Grid.Col < Grid.Cols - 1 Then
            'Grid.Col = Grid.Col + 1
        Else
            If txtEdit.SelStart = Len(txtEdit) And Grid.Row < Grid.Rows - 1 Then
                'Grid.Row = Grid.Row + 1
                'Grid.Col = 1
            End If
        End If
    Case 40     'Dn Arrow
        If Grid.Row < Grid.Rows - 1 Then
            Grid.Row = Grid.Row + 1
        End If
End Select
End Sub

Sub txtKeyPress(txtEdit As Control, KeyAscii As Integer, Grid As MSFlexGrid)
Dim pos%, l$, r$
Select Case KeyAscii
    Case 13
        'sometimes this puts chr$(3) at the end
        'KeyAscii = 0
        Grid.Text = Replace(txtEdit.Text, Chr$(3), "", 1, , vbTextCompare)
        Dirty = 1
        txtEdit.Visible = False
        Rowp = Rowp + 1
        'grid.CellBackColor = WHITE
        If Rowp < Grid.Rows Then Grid.Row = Rowp
        Grid.SetFocus
    Case 8                      'BkSpc - split string @ cursor
        'pos% = txtEdit.SelStart - 1 'where is the cursor?
        'If pos% >= 0 Then
        '    l$ = Left$(grid, pos%)       'left of cursor
        '    R$ = Right$(grid, Len(grid) - pos% - 1) 'right of cursor
        '    grid.Text = l$ + R$          'depleted string into grid
        'End If
    Case 27, 37 To 40
        Grid = txtEdit        'or it's going to look funny
        Dirty = 1
        txtEdit.Visible = False
    Case Else
        Grid = txtEdit + Chr(KeyAscii)
        Dirty = 1
End Select
End Sub

Function AddRows(Grid As MSFlexGrid, InsRows As Integer)
On Error GoTo Oops
Dim LastGridRowNum As Integer
Dim NewRowNum As Integer
Dim GridAdd$
Dim gC As Integer
Dim OldRows As Integer
Dim LastRow As Integer
Dim j As Integer
If InsRows = 0 Then GoTo Exit_AddRows
Screen.MousePointer = 11
Grid.Redraw = False
InitGrid = 0
Grid.Row = GSR
Grid.Col = 1
If GSR > 0 Then
    LastGridRowNum = Val(Grid.TextMatrix(GSR - 1, 0))
End If
For i = 0 To InsRows - 1
    NewRowNum = GSR + i
    GridAdd$ = Trim(Str$(LastGridRowNum + i + 1))
    For gC = 1 To Grid.Cols - 1
        GridAdd$ = GridAdd$ & vbTab
    Next gC
    If NewRowNum > 0 Then
        Grid.AddItem GridAdd$, NewRowNum
    End If
Next i
For j = NewRowNum + 1 To Grid.Rows - 2
    gCol = 0
    gRow = j
    Grid.TextMatrix(gRow, gCol) = Trim$(Str$(LastGridRowNum + j - NewRowNum + 1))
Next j
GoTo Exit_AddRows
'-------------------------------------------------------
Grid.RowSel = Grid.Rows - 1
Grid.ColSel = Grid.Cols - 1
'grid.Redraw = True
TempClip$ = Grid.Clip
NewClip$ = ""
RowClip$ = ""
For i = 1 To Grid.Cols
    RowClip$ = RowClip$ & Chr$(9)
Next i
OldRows = Grid.Rows
Grid.Rows = OldRows + InsRows
LastRow = OldRows - 2 'bottom row is empty
For i = 0 To InsRows 'was 1
    gRow = LastRow + i
    gCol = 0
    Grid.TextMatrix(gRow, gCol) = Trim$(Str$(gRow - 1))
    NewClip$ = NewClip$ & RowClip$ & vbCr
Next i
Grid.Clip = NewClip$ ' ""
Grid.Row = GSR + InsRows
Grid.Col = 1
Grid.RowSel = Grid.Rows - 1
Grid.ColSel = Grid.Cols - 1
Grid.Clip = TempClip$
Grid.Row = GSR
Grid.RowSel = GSR
On Error GoTo gsrerr
Grid.TopRow = GSR
InitGrid = 1
GoTo Exit_AddRows
gsrerr:
If Err = 30009 Then GSR = GSR - 1: Resume
GoTo Exit_AddRows
Oops:
If Err = 32755 Then
    GoTo Exit_AddRows
End If
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine AddRows "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in AddRows"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_AddRows:
End Function

Sub SetTextBox(txtEdit As TextBox, Grid As MSFlexGrid)
'put textbox over cell
With txtEdit
    .Top = (Grid.Top + Grid.CellTop) '- 50
    .Left = Grid.Left + Grid.CellLeft
    .Width = Grid.CellWidth
    .Height = Grid.CellHeight
    Grid.Text = Replace(Grid.Text, Chr$(3), "", 1, , vbTextCompare)
    .Text = Grid.Text
    .Visible = True
    .SelStart = 0
    .SelLength = Len(.Text)
    '.SelStart = Len(.Text)
    '.Width = Me.TextWidth(.Text & " ")
    .SetFocus
    .ZOrder 0
End With
DoEvents
End Sub
'**************************************
' Name: Auto Size Grid Column
' Description:Automatically Size a Column of a MSFlexGrid.
' By: Chris Dostal
'
' Inputs:Grid as MSFlexGrid, inCol as Column you want to autosize.
'
'This code is copyrighted and has limited warranties.Please see
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=51189&lngWId=1
'for details.
'**************************************

Function AutoSizeGridCol(Grid As MSFlexGrid, InCol As Long, Optional SetU As Integer, Optional SetCaps As Integer)
'Will analyse TEXTMatrix(y,incol) and figure out the longest string.
'From this will do a calculation (includes header)
'250 for first char and u for each otherone..
Dim U As Integer
Dim y As Integer
Dim t As Integer
Dim Caps As Integer
Dim MaxLen As Integer
Dim Unit As Integer
Dim StringIn$
Dim buf$
U = 88
Caps = 110
'Override with optional command parameters..
If SetU > 0 Then U = SetU
If SetCaps > 0 Then Caps = SetCaps
MaxLen = 0 'flag
For y = 0 To Grid.Rows - 1
    StringIn$ = Grid.TextMatrix(y, InCol)
    Unit = 0 'reset
    For t = 1 To Len(StringIn$)
        buf$ = Mid(StringIn$, t, 1)
        Select Case Asc(buf$)
            Case Is >= 97, Is <= 122
                Unit = Unit + U
            Case Is <= 96, Is >= 123
                Unit = Unit + Caps
        End Select
    Next t
    If Unit > MaxLen Then MaxLen = Unit
Next y
If MaxLen > 1 Then
    AutoSizeGridCol = MaxLen
    Grid.ColWidth(InCol) = 250 + AutoSizeGridCol
    Exit Function
End If
AutoSizeGridCol = 250
Grid.ColWidth(InCol) = AutoSizeGridCol
End Function

'**************************************
' Name: Flexgrid autoresize auto size columns
' Description:A routine to autoresize columns in the MSHflexgrid.
'There is a previous example of PSC but this one is much
'faster and doesn't make the grid flicker.
' By: blackbc
'
'This code is copyrighted and has limited warranties.Please see
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39225&lngWId=1
'for details.
'**************************************
Public Sub SetGridColumnWidth(Grid As MSFlexGrid)
'params:ms flexgrid control
'purpose:sets the column widths to the
'lengths of the longest string in the column
'requirements: the grid must have the same
'font as the underlying form
Dim InnerLoopCount As Long
Dim OuterLoopCount As Long
Dim lngLongestLen As Long
Dim sLongestString As String
Dim lngColWidth As Long
Dim szCellText As String
With Grid
    For OuterLoopCount = 0 To .Cols - 1
        sLongestString = ""
        lngLongestLen = 0
        For InnerLoopCount = 0 To .Rows - 1
            szCellText = .TextMatrix(InnerLoopCount, OuterLoopCount)
            If Len(szCellText) > lngLongestLen Then
                lngLongestLen = Len(szCellText)
                sLongestString = szCellText
            End If
        Next
        lngColWidth = .Parent.TextWidth(sLongestString)
        'add 100 for more readable spreadsheet
        .ColWidth(OuterLoopCount) = lngColWidth + 200
    Next
End With
End Sub
