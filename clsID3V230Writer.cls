VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3V230Writer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private fstrSongTitle As String
Private fstrArtist As String
Private fstrAlbum As String
Private fstrYear As String
Private fstrComment As String
Private fstrTrack As String
Private fstrGenre As String
Private fstrComposer As String
Private fstrOriginalArtist As String
Private fstrCopyright As String
Private fstrURL As String
Private fstrEncodedBy As String
Private fstrLyrics As String
Private fstrSyncLyrics As String
Private fstrGraphic As String
Private fstrGraphicType As String
Private fstrGraphicTitle As String
Private fstrSubTitle As String
Private fstrLanguage As String

Public Property Let SubTitle(ByVal strSubTitle As String)
fstrSubTitle = strSubTitle
End Property

Public Property Get SubTitle() As String
SubTitle = fstrSubTitle
End Property

Public Property Let Language(ByVal strLanguage As String)
fstrLanguage = strLanguage
End Property

Public Property Get Language() As String
Language = fstrLanguage
End Property

Public Property Let Graphic(ByVal strGraphic As String)
fstrGraphic = strGraphic
End Property

Public Property Get Graphic() As String
Graphic = fstrGraphic
End Property

Public Property Let GraphicType(ByVal strGraphicType As String)
fstrGraphicType = strGraphicType
End Property

Public Property Get GraphicType() As String
GraphicType = fstrGraphicType
End Property

Public Property Let GraphicTitle(ByVal strGraphicTitle As String)
fstrGraphicTitle = strGraphicTitle
End Property

Public Property Get GraphicTitle() As String
GraphicTitle = fstrGraphicTitle
End Property

Public Property Let Lyrics(ByVal strLyrics As String)
fstrLyrics = strLyrics
End Property

Public Property Get Lyrics() As String
Lyrics = fstrLyrics
End Property

Public Property Let SyncLyrics(ByVal strSyncLyrics As String)
fstrSyncLyrics = strSyncLyrics
End Property

Public Property Get SyncLyrics() As String
SyncLyrics = fstrSyncLyrics
End Property

Public Property Let Composer(ByVal strComposer As String)
fstrComposer = strComposer
End Property

Public Property Get Composer() As String
Composer = fstrComposer
End Property

Public Property Let OriginalArtist(ByVal strOriginalArtist As String)
fstrOriginalArtist = strOriginalArtist
End Property

Public Property Get OriginalArtist() As String
OriginalArtist = fstrOriginalArtist
End Property

Public Property Let Copyright(ByVal strCopyright As String)
fstrCopyright = strCopyright
End Property

Public Property Get Copyright() As String
Copyright = fstrCopyright
End Property

Public Property Let URL(ByVal strURl As String)
fstrURL = strURl
End Property

Public Property Get URL() As String
URL = fstrURL
End Property

Public Property Let EncodedBy(ByVal strEncodedBy As String)
fstrEncodedBy = strEncodedBy
End Property

Public Property Get EncodedBy() As String
EncodedBy = fstrEncodedBy
End Property

Public Property Let SongTitle(ByVal strSongTitle As String)
fstrSongTitle = strSongTitle
End Property

Public Property Get SongTitle() As String
SongTitle = fstrSongTitle
End Property

Public Property Let Artist(ByVal strArtist As String)
fstrArtist = strArtist
End Property

Public Property Get Artist() As String
Artist = fstrArtist
End Property

Public Property Let Album(ByVal strAlbum As String)
fstrAlbum = strAlbum
End Property

Public Property Get Album() As String
Album = fstrAlbum
End Property

Public Property Let Year(ByVal strYear As String)
fstrYear = strYear
End Property

Public Property Get Year() As String
Year = fstrYear
End Property

Public Property Let Comment(ByVal strComment As String)
fstrComment = strComment
End Property

Public Property Get Comment() As String
Comment = fstrComment
End Property

Public Property Let Track(ByVal strTrack As String)
fstrTrack = strTrack
End Property

Public Property Get Track() As String
Track = fstrTrack
End Property

Public Property Let Genre(ByVal strGenre As String)
fstrGenre = strGenre
End Property

Public Property Get Genre() As String
Genre = fstrGenre
End Property

Public Sub RemoveID3V230Tag(ByVal strFileName As String)
Dim intFile As Integer
Dim intFile2 As Integer
Dim ID3 As String * 3
Dim bytTemp() As Byte
Dim lngTotalSize As Long
Dim bytByte() As Byte
Dim lngSizeLeft As Long
Dim lngBuffer As Long
On Error GoTo ErrHan
If FileExists(strFileName) = False Then Exit Sub
intFile = FreeFile
Open strFileName For Binary Access Read Shared As intFile
Get #intFile, , ID3
If ID3 = "ID3" Then
    ReDim bytTemp(1 To 7)
    Get #intFile, , bytTemp
    lngTotalSize = LeftShift(bytTemp(4), 21) + LeftShift(bytTemp(5), 14) + LeftShift(bytTemp(6), 7) + bytTemp(7) + 10
End If
Close
If lngTotalSize <> 0 Then
    intFile = FreeFile
    ReDim bytTemp(1 To lngTotalSize)
    Open strFileName For Binary Access Read Shared As intFile
    Get #intFile, , bytTemp
    intFile2 = FreeFile
    lngSizeLeft = LOF(intFile) - lngTotalSize
    lngBuffer = 2097152
    Open strFileName & ".new" For Binary Access Write As intFile2
    Do Until (EOF(intFile)) Or (lngSizeLeft = 0)
        If (lngSizeLeft - lngBuffer) >= 0 Then
            ReDim bytByte(1 To lngBuffer)
        Else
            lngBuffer = lngSizeLeft
            ReDim bytByte(1 To lngBuffer)
        End If
        Get #intFile, , bytByte
        Put #intFile2, , bytByte
        lngSizeLeft = lngSizeLeft - lngBuffer
        DoEvents
    Loop
    Close intFile2
    Close intFile
    Kill strFileName
    Name strFileName & ".new" As strFileName
End If
Exit Sub
ErrHan:
End Sub

Public Sub WriteID3V230Tag(ByVal strFileName As String)
Dim lngTotalTagSize As Long
Dim intNewFile As Integer
Dim bytSyncSafe() As Byte
Dim lngFileLength As Long
Dim lngBytesToRead As Long
Dim lngBuffer As Long
Dim bytBytes() As Byte
Dim intOldFile As Integer
Dim strURLS() As String
Dim strGraphics() As String
Dim strGraphicTypes() As String
Dim strGraphicTitles() As String
Dim lngCounter As Long
Dim lngSyncLyricsSize As Long
On Error GoTo ErrHan
RemoveID3V230Tag strFileName
If Comment <> "" Then Comment = IIf(Len(Language) <> 3, "eng", Language) & Chr(0) & Comment
If Len(SongTitle) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(SongTitle) + 11
End If
If Len(Artist) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Artist) + 11
End If
If Len(Album) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Album) + 11
End If
If Len(Year) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Year) + 11
End If
If Len(Comment) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Comment) + 11
End If
If Len(Track) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Track) + 11
End If
If Len(Genre) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Genre) + 11
End If
If Len(URL) <> 0 Then
    If InStr(URL, "|") <> 0 Then
        strURLS = Split(URL, "|")
        For lngCounter = 0 To UBound(strURLS)
            lngTotalTagSize = lngTotalTagSize + Len(strURLS(lngCounter)) + 12
        Next
    Else
        ReDim strURLS(0)
        strURLS(0) = URL
        lngTotalTagSize = lngTotalTagSize + Len(strURLS(0)) + 12
    End If
End If
If Len(Composer) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Composer) + 11
End If
If Len(OriginalArtist) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(OriginalArtist) + 11
End If
If Len(Copyright) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Copyright) + 11
End If
If Len(EncodedBy) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(EncodedBy) + 11
End If
If Len(SubTitle) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(SubTitle) + 11
End If
If Len(Language) <> 0 Then
    lngTotalTagSize = lngTotalTagSize + Len(Language) + 11
End If
If Len(Lyrics) <> 0 Then
    Lyrics = "eng" & Chr(0) & Lyrics
    lngTotalTagSize = lngTotalTagSize + Len(Lyrics) + 11
End If
If Len(Graphic) <> 0 Then
    If InStr(Graphic, "|") <> 0 Then
        strGraphics = Split(Graphic, "|")
        strGraphicTypes = Split(GraphicType, "|")
        strGraphicTitles = Split(GraphicTitle, "|")
        For lngCounter = 0 To UBound(strGraphics)
            If FileExists(strGraphics(lngCounter)) = True Then lngTotalTagSize = lngTotalTagSize + FileLen(strGraphics(lngCounter)) + 20 + Len(GetImageType(strGraphics(lngCounter))) + Len(strGraphicTitles(lngCounter))
        Next
    Else
        ReDim strGraphics(0)
        strGraphics(0) = Graphic
        ReDim strGraphicTypes(0)
        strGraphicTypes(0) = GraphicType
        ReDim strGraphicTitles(0)
        strGraphicTitles(0) = GraphicTitle
        If FileExists(strGraphics(0)) = True Then lngTotalTagSize = lngTotalTagSize + FileLen(strGraphics(0)) + 20 + Len(GetImageType(strGraphics(0))) + Len(strGraphicTitles(0))
    End If
End If
If Len(SyncLyrics) <> 0 Then
    lngSyncLyricsSize = SizeOfSyncLyrics(SyncLyrics)
    lngTotalTagSize = lngTotalTagSize + lngSyncLyricsSize + 8
End If
ReDim bytSyncSafe(3)
CreateSyncSafeInteger bytSyncSafe, lngTotalTagSize
If lngTotalTagSize <> 0 Then
    lngTotalTagSize = lngTotalTagSize + 10
    intNewFile = FreeFile
    Open strFileName & ".mp3" For Binary Access Write As intNewFile
    Put #intNewFile, , "ID3"
    Put #intNewFile, , CByte(3)
    Put #intNewFile, , CByte(0)
    Put #intNewFile, , CByte(0)
    Put #intNewFile, , bytSyncSafe
    If Len(SongTitle) <> 0 Then
        Put #intNewFile, , "TIT2"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(SongTitle) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , SongTitle
    End If
    If Len(Artist) <> 0 Then
        Put #intNewFile, , "TPE1"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Artist) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Artist
    End If
    If Len(Album) <> 0 Then
        Put #intNewFile, , "TALB"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Album) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Album
    End If
    If Len(Year) <> 0 Then
        Put #intNewFile, , "TYER"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Year) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Year
    End If
    If Len(Track) <> 0 Then
        Put #intNewFile, , "TRCK"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Track) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Track
    End If
    If Len(Genre) <> 0 Then
        Put #intNewFile, , "TCON"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Genre) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Genre
    End If
    If Len(Comment) <> 0 Then
        Put #intNewFile, , "COMM"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Comment) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Comment
    End If
    If Len(Composer) <> 0 Then
        Put #intNewFile, , "TCOM"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Composer) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Composer
    End If
    If Len(OriginalArtist) <> 0 Then
        Put #intNewFile, , "TOPE"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(OriginalArtist) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , OriginalArtist
    End If
    If Len(Language) <> 0 Then
        Put #intNewFile, , "TLAN"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Language) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Language
    End If
    If Len(SubTitle) <> 0 Then
        Put #intNewFile, , "TIT3"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(SubTitle) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , SubTitle
    End If
    If Len(Copyright) <> 0 Then
        Put #intNewFile, , "TCOP"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(Copyright) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Copyright
    End If
    If Len(EncodedBy) <> 0 Then
        Put #intNewFile, , "TENC"
        ReDim bytSyncSafe(3)
        CreateSyncSafeInteger bytSyncSafe, Len(EncodedBy) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , EncodedBy
    End If
    If Len(URL) <> 0 Then
        For lngCounter = 0 To UBound(strURLS)
            Put #intNewFile, , "WXXX"
            ReDim bytSyncSafe(3)
            CreateSyncSafeInteger bytSyncSafe, Len(strURLS(lngCounter)) + 2
            Put #intNewFile, , bytSyncSafe
            Put #intNewFile, , CByte(0)
            Put #intNewFile, , CByte(0)
            Put #intNewFile, , CByte(0)
            Put #intNewFile, , CByte(0)
            Put #intNewFile, , strURLS(lngCounter)
        Next
    End If
    If Len(Lyrics) <> 0 Then
        Put #intNewFile, , "USLT"
        ReDim bytSyncSafe(3)
        CreateSyncInteger bytSyncSafe, Len(Lyrics) + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , Lyrics
    End If
    If lngSyncLyricsSize > 0 Then
        Put #intNewFile, , "SYLT"
        CreateSyncInteger bytSyncSafe, lngSyncLyricsSize + 1
        Put #intNewFile, , bytSyncSafe
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , CByte(0)
        Put #intNewFile, , "eng"
        Put #intNewFile, , CByte(2)
        Put #intNewFile, , CByte(1)
        Put #intNewFile, , CByte(0)
        WriteSyncLyrics intNewFile
    End If
    If Len(Graphic) > 0 Then
        For lngCounter = 0 To UBound(strGraphics)
            If FileExists(strGraphics(lngCounter)) = True Then
                Put #intNewFile, , "APIC"
                ReDim bytSyncSafe(3)
                CreateSyncInteger bytSyncSafe, FileLen(strGraphics(lngCounter)) + 10 + Len(GetImageType(strGraphics(lngCounter))) + Len(strGraphicTitles(lngCounter))
                Put #intNewFile, , bytSyncSafe
                Put #intNewFile, , CByte(0)
                Put #intNewFile, , CByte(0)
                Put #intNewFile, , CByte(0)
                Put #intNewFile, , "image/" & GetImageType(strGraphics(lngCounter))
                Put #intNewFile, , CByte(0)
                Put #intNewFile, , CByte(strGraphicTypes(lngCounter))
                If strGraphicTitles(lngCounter) <> "" Then Put #intNewFile, , strGraphicTitles(lngCounter)
                Put #intNewFile, , CByte(0)
                InsertFile intNewFile, strGraphics(lngCounter)
            End If
        Next
    End If
    InsertFile intNewFile, strFileName
    Close intNewFile
    Kill strFileName
    Name strFileName & ".mp3" As strFileName
End If
Exit Sub
ErrHan:
End Sub

Private Sub InsertFile(ByVal intFile As Integer, ByVal strFile As String)
Dim intRead As Integer
Dim lngBytesToRead As Long
Dim lngBuffer As Long
Dim bytBytes() As Byte
On Error GoTo ErrHan
intRead = FreeFile
Open strFile For Binary Access Read Shared As intRead
lngBytesToRead = LOF(intRead)
lngBuffer = 2097152
ReDim bytBytes(1 To lngBuffer)
Do
    If lngBytesToRead < lngBuffer Then
        lngBuffer = lngBytesToRead
        ReDim bytBytes(1 To lngBuffer)
    End If
    Get #intRead, , bytBytes
    Put #intFile, , bytBytes
    lngBytesToRead = lngBytesToRead - lngBuffer
    DoEvents
Loop Until (lngBytesToRead <= 0) Or (EOF(intRead))
Close intRead
Exit Sub
ErrHan:
End Sub

Private Sub CreateSyncInteger(ByRef bytBytes() As Byte, ByVal lngSize As Long)
Dim strHex As String
Dim lngPad As Long
Dim lngCounter As Long
Dim lngByte As Long
On Error GoTo ErrHan
strHex = Hex(lngSize)
lngPad = 8 - Len(strHex)
If lngPad > 0 Then strHex = String(lngPad, "0") & strHex
lngByte = 0
For lngCounter = 1 To Len(strHex) Step 2
    bytBytes(lngByte) = CByte("&H" & Mid(strHex, lngCounter, 2))
    lngByte = lngByte + 1
Next
Exit Sub
ErrHan:
End Sub

Private Function GetImageType(ByVal strFile As String) As String
Dim strExt As String
On Error GoTo ErrHan
strExt = LCase(Mid(strFile, InStrRev(strFile, ".") + 1))
Select Case strExt
    Case "jpg", "jpeg"
        GetImageType = "jpg"
    Case Else
        GetImageType = ""
End Select
Exit Function
ErrHan:
End Function

Private Function LeftShift(ByVal intNumber As Integer, ByVal bytAmount As Byte) As Long
Dim strTemp As String
If bytAmount = 0 Then
    LeftShift = intNumber
Else
    strTemp = NumberToBinary(intNumber)
    strTemp = strTemp & String(bytAmount, "0")
    LeftShift = BinaryToLong(strTemp)
End If
End Function

Private Function BinaryToLong(ByVal strBinary As String) As Long
Dim lngTotal As Long
Dim lngFactor As Long
Dim lngCounter As Long
Dim bytBit As Byte
lngFactor = 0
For lngCounter = Len(strBinary) To 1 Step -1
    bytBit = CByte(Mid(strBinary, lngCounter, 1))
    lngTotal = lngTotal + (bytBit * (2 ^ lngFactor))
    lngFactor = lngFactor + 1
Next
BinaryToLong = lngTotal
End Function

Private Function NumberToBinary(ByVal intNumber As Integer) As String
Dim lngNumber As Long
Dim strTemp As String
Dim strPad As String
On Error GoTo ErrHan
lngNumber = intNumber
Do
    If (lngNumber Mod 2) = 0 Then
        strTemp = "0" & strTemp
    Else
        strTemp = "1" & strTemp
    End If
    lngNumber = lngNumber \ 2
    DoEvents
Loop Until lngNumber = 0
If Len(strTemp) < 8 Then strPad = String(8 - Len(strTemp), "0")
NumberToBinary = strPad & strTemp
Exit Function
ErrHan:
NumberToBinary = ""
End Function

Private Sub CreateSyncSafeInteger(ByRef SyncsafeInt() As Byte, ByVal lngValue As Long)
Dim intCounter As Integer
Dim intCurrentByte As Integer
Dim intCurrentBit As Integer
intCurrentByte = 3
For intCounter = 0 To 27
    If lngValue And 2 ^ intCounter Then
        SetBit intCurrentBit, SyncsafeInt(intCurrentByte), True
    End If
    intCurrentBit = intCurrentBit + 1
    If intCurrentBit Mod 7 = 0 Then
        intCurrentByte = intCurrentByte - 1
        intCurrentBit = 0
    End If
Next
End Sub

Private Sub SetBit(ByVal intBit As Integer, ByRef bytSetByte As Byte, ByVal bolBitState As Boolean)
Dim bytPattern As Byte
If bolBitState = True Then
    bytPattern = 2 ^ intBit
    bytSetByte = bytSetByte Or bytPattern
Else
    bytPattern = 255 - (2 ^ intBit)
    bytSetByte = bytSetByte And bytPattern
End If
End Sub

Private Function FileExists(ByVal strFileName As String) As Boolean
Dim intFile As Integer
On Error GoTo ErrHan
intFile = FreeFile
Open strFileName For Input As intFile
Close intFile
FileExists = True
Exit Function
ErrHan:
FileExists = False
End Function

Public Function GetGraphicType(ByVal strType As String) As Byte
Select Case LCase(strType)
    Case "32x32 icon"
        GetGraphicType = 1
    Case "other icon"
        GetGraphicType = 2
    Case "front cover"
        GetGraphicType = 3
    Case "back cover"
        GetGraphicType = 4
    Case "leaflet page"
        GetGraphicType = 5
    Case "media"
        GetGraphicType = 6
    Case "lead performer"
        GetGraphicType = 7
    Case "artist"
        GetGraphicType = 8
    Case "conductor"
        GetGraphicType = 9
    Case "band"
        GetGraphicType = 10
    Case "composer"
        GetGraphicType = 11
    Case "lyricist"
        GetGraphicType = 12
    Case "recording location"
        GetGraphicType = 13
    Case "during recording"
        GetGraphicType = 14
    Case "during performance"
        GetGraphicType = 15
    Case "movie capture"
        GetGraphicType = 16
    Case "a bright colored fish"
        GetGraphicType = 17
    Case "illustration"
        GetGraphicType = 18
    Case "band logo"
        GetGraphicType = 19
    Case "publisher logo"
        GetGraphicType = 20
    Case Else
        GetGraphicType = 0
End Select
End Function

Private Function SizeOfSyncLyrics(ByVal strSyncLyrics As String) As Long
Dim strLines() As String
Dim lngCounter As Long
Dim lngInstr As Long
Dim strTime As String
Dim strLine As String
Dim lngSize As Long
On Local Error GoTo ErrHan
strLines = Split(strSyncLyrics, vbCrLf)
For lngCounter = 0 To UBound(strLines)
    lngInstr = InStr(strLines(lngCounter), "]")
    If lngInstr <> 0 Then
        strTime = Left(strLines(lngCounter), lngInstr)
        strLine = Right(strLines(lngCounter), Len(strLines(lngCounter)) - Len(strTime))
        strTime = Mid(strTime, 2, Len(strTime) - 2)
        strTime = GetMilliseconds(strTime)
        If IsNumeric(strTime) = False Then
            Exit Function
        Else
            lngSize = lngSize + 4 + Len(strLine) + 1 + 1
        End If
    End If
Next
SizeOfSyncLyrics = lngSize + 6
Exit Function
ErrHan:
End Function

Private Sub WriteSyncLyrics(ByVal intFileNumber As Integer)
Dim strLines() As String
Dim lngCounter As Long
Dim lngInstr As Long
Dim strTime As String
Dim lngTime As Long
Dim strLine As String
Dim bytBytes() As Byte
On Local Error GoTo ErrHan
strLines = Split(SyncLyrics, vbCrLf)
For lngCounter = 0 To UBound(strLines)
    lngInstr = InStr(strLines(lngCounter), "]")
    If lngInstr <> 0 Then
        strTime = Left(strLines(lngCounter), lngInstr)
        strLine = Right(strLines(lngCounter), Len(strLines(lngCounter)) - Len(strTime))
        strTime = Mid(strTime, 2, Len(strTime) - 2)
        strTime = GetMilliseconds(strTime)
        If IsNumeric(strTime) = False Then
            lngTime = 0
        Else
            lngTime = CLng(strTime)
            ReDim bytBytes(0 To 3)
            CreateSyncInteger bytBytes, lngTime
        End If
    End If
    Put #intFileNumber, , strLine
    Put #intFileNumber, , CByte(0)
    Put #intFileNumber, , bytBytes
    Put #intFileNumber, , CByte(10)
Next
Exit Sub
ErrHan:
MsgBox Err.Description
End Sub

Private Function GetMilliseconds(ByVal strTime As String) As Long
Dim strParts() As String
On Local Error GoTo ErrHan
strParts = Split(strTime, ":")
GetMilliseconds = ((CLng(strParts(0)) * 60) + CLng(strParts(1))) * 1000
Exit Function
ErrHan:
End Function
