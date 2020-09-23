VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3V230Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private fbolTagPresent As Boolean
Private fstrSongTitle As String
Private fstrArtist As String
Private fstrAlbum As String
Private fstrYear As String
Private fstrComment As String
Private fbytTrack As Byte
Private fbytTracks As Byte
Private fstrGenre As String
Private fstrComposer As String
Private fstrOriginalArtist As String
Private fstrCopyright As String
Private fstrURL As String
Private fstrEncodedBy As String
Private fstrLyrics As String
Private fstrSyncLyrics As String
Private fstrGraphicStartPos As String
Private fstrGraphicSize As String
Private fstrGraphicExtended As String
Private fstrGraphicTitle As String
Private fbolExperimentalTag As Boolean
Private fbolExtendedHeader As Boolean
Private fbolCRC As Boolean
Private fdblCRCValue As Double
Private fstrSubTitle As String
Private fstrLanguage As String
Private fstrUnknown As String
Private Type ID3Header
    strID3 As String * 3
    bytVersionMajor As Byte
    bytVersionMinor As Byte
    bytFlags As Byte
    bytSize(0 To 3) As Byte
End Type
Private Type ID3Frame
    strID As String * 4
    bytSize(0 To 3) As Byte
    bytFlags(0 To 1) As Byte
End Type

Public Property Let Unknown(ByVal strUnknown As String)
fstrUnknown = strUnknown
End Property

Public Property Get Unknown() As String
Unknown = fstrUnknown
End Property

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

Public Property Let CRCValue(ByVal dblCRCValue As Double)
fdblCRCValue = dblCRCValue
End Property

Public Property Get CRCValue() As Double
CRCValue = fdblCRCValue
End Property

Public Property Let CRC(ByVal bolCRC As Boolean)
fbolCRC = bolCRC
End Property

Public Property Get CRC() As Boolean
CRC = fbolCRC
End Property

Public Property Let ExtendedHeader(ByVal bolExtendedHeader As Boolean)
fbolExtendedHeader = bolExtendedHeader
End Property

Public Property Get ExtendedHeader() As Boolean
ExtendedHeader = fbolExtendedHeader
End Property

Public Property Let ExperimentalTag(ByVal bolExperimentalTag As Boolean)
fbolExperimentalTag = bolExperimentalTag
End Property

Public Property Get ExperimentalTag() As Boolean
ExperimentalTag = fbolExperimentalTag
End Property

Public Property Let GraphicTitle(ByVal strGraphicTitle As String)
fstrGraphicTitle = strGraphicTitle
End Property

Public Property Get GraphicTitle() As String
GraphicTitle = fstrGraphicTitle
End Property

Public Property Let GraphicSize(ByVal strGraphicSize As String)
fstrGraphicSize = strGraphicSize
End Property

Public Property Get GraphicSize() As String
GraphicSize = fstrGraphicSize
End Property

Public Property Let GraphicStartPos(ByVal strGraphicStartPos As String)
fstrGraphicStartPos = strGraphicStartPos
End Property

Public Property Get GraphicStartPos() As String
GraphicStartPos = fstrGraphicStartPos
End Property

Public Property Let GraphicExtended(ByVal strGraphicExtended As String)
fstrGraphicExtended = strGraphicExtended
End Property

Public Property Get GraphicExtended() As String
GraphicExtended = fstrGraphicExtended
End Property

Public Property Let SyncLyrics(ByVal strSyncLyrics As String)
fstrSyncLyrics = strSyncLyrics
End Property

Public Property Get SyncLyrics() As String
SyncLyrics = fstrSyncLyrics
End Property

Public Property Let Lyrics(ByVal strLyrics As String)
fstrLyrics = strLyrics
End Property

Public Property Get Lyrics() As String
Lyrics = fstrLyrics
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

Public Property Let TagPresent(ByVal bolTagPresent As Boolean)
fbolTagPresent = bolTagPresent
End Property

Public Property Get TagPresent() As Boolean
TagPresent = fbolTagPresent
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

Public Property Let Track(ByVal bytTrack As Byte)
fbytTrack = bytTrack
End Property

Public Property Get Track() As Byte
Track = fbytTrack
End Property

Public Property Let Tracks(ByVal bytTracks As Byte)
fbytTracks = bytTracks
End Property

Public Property Get Tracks() As Byte
Tracks = fbytTracks
End Property

Public Property Let Genre(ByVal strGenre As String)
fstrGenre = strGenre
End Property

Public Property Get Genre() As String
Genre = fstrGenre
End Property

Public Sub ReadID3V230Tag(ByVal strFileName As String, Optional ByVal bolAutoReadUnknown As Boolean = False, Optional ByVal bolDetermineTagPresent As Boolean = False)
Dim intFile As Integer
Dim ID3 As ID3Header
Dim bolUnsync As Boolean
Dim bolExtended As Boolean
Dim bolExperimental As Boolean
Dim dblTotalTagSize As Double
Dim dblBytesRead As Double
Dim bytExtendedSize() As Byte
Dim bytExtendedFlags() As Byte
Dim bytExtendedPadding() As Byte
Dim bytExtendedCRC() As Byte
Dim dblExtendedHeaderSize As Double
Dim Frame As ID3Frame
Dim dblFrameSize As Double
Dim bytFrameData() As Byte
Dim bolTagAlterDiscard As Boolean
Dim bolFileAlterDiscard As Boolean
Dim bolReadOnly As Boolean
Dim bolCompressed As Boolean
Dim bolEncrypted As Boolean
Dim bolGrouped As Boolean
Dim strData As String
Dim strTrackInfo() As String
Dim bytTextEncoding As Byte
Dim strCommentInfo() As String
Dim strURLInfo() As String
Dim strLyricsInfo() As String
Dim bytInternalImageType As Byte
Dim lngResult As Long
Dim strDescriptor As String
On Error GoTo ErrHan
If FileExists(strFileName) = False Then Exit Sub
ClearProperties
intFile = FreeFile
Open strFileName For Binary Access Read Shared As intFile
Get #intFile, , ID3
'If this isn't the right version, stop reading
If (ID3.strID3 <> "ID3") Or (ID3.bytVersionMajor <> 3) Then GoTo ErrHan
If bolDetermineTagPresent = True Then
    TagPresent = True
    Close intFile
    Exit Sub
End If
dblBytesRead = Len(ID3)
'Check the ID3 Flags Bit
GetFlagInformation ID3.bytFlags, bolUnsync, bolExtended, bolExperimental
dblTotalTagSize = DecodeUnsync(ID3.bytSize)
'We can't read Unsynced frames...so stop trying
If bolUnsync = True Then
    MsgBox "This ID3 Tag is Unsyncronized.  This application is not equipped to read this kind of ID3 Tag.", vbOKOnly + vbInformation
    GoTo ErrHan
End If
ExperimentalTag = bolExperimental
ExtendedHeader = bolExtended
'Read in the extended header
If ExtendedHeader = True Then
    ReDim bytExtendedSize(4)
    ReDim bytExtendedFlags(2)
    ReDim bytExtendedPadding(4)
    Get #intFile, , bytExtendedSize
    Get #intFile, , bytExtendedFlags
    Get #intFile, , bytExtendedPadding
    dblExtendedHeaderSize = DecodeUnsync(bytExtendedSize)
    'If there's a CRC, read it too
    If dblExtendedHeaderSize = 10 Then
        ReDim bytExtendedCRC(4)
        Get #intFile, , bytExtendedCRC
        CRC = True
        CRCValue = DecodeUnsync(bytExtendedCRC)
    End If
    dblBytesRead = dblBytesRead + dblExtendedHeaderSize + UBound(bytExtendedSize)
End If
'Loop through the frames
Do
NextFrame:
    DoEvents
    Get #intFile, , Frame
    
    If Frame.strID = Chr(0) & Chr(0) & Chr(0) & Chr(0) Then Exit Do
    dblBytesRead = dblBytesRead + Len(Frame)
    'According to the documentation, the frame size is the size of the frame,
    'excluding the 10 byte header size.  Note also, that the size given is 1
    'byte more that the actual size...
    DecodeFrameFlags Frame.bytFlags, bolTagAlterDiscard, bolFileAlterDiscard, bolReadOnly, bolCompressed, bolEncrypted, bolGrouped
    'Not sure what this is about, but it seems like the size of the lyrics tag
    'is not saved as an Unsynched Integer...strange...
    Select Case Frame.strID
        Case "USLT", "SYLT"
            dblFrameSize = DecodeSync(Frame.bytSize) - 1
        Case "APIC"
            dblFrameSize = DecodeSync(Frame.bytSize) - 1
            If GraphicStartPos = "" Then
                GraphicStartPos = dblBytesRead - 10
                GraphicSize = dblFrameSize + 1
            Else
                GraphicStartPos = GraphicStartPos & "|" & dblBytesRead - 10
                GraphicSize = GraphicSize & "|" & dblFrameSize + 1
            End If
        Case "WXXX"
            dblFrameSize = DecodeUnsync(Frame.bytSize) - 1
        Case "TIT2", "TPE1", "TALB", "TYER", "COMM", "TCON", "TRCK", _
            "TENC", "TCOM", "TOPE", "TCOP", "TIT3", "TLAN"
            dblFrameSize = DecodeUnsync(Frame.bytSize) - 1
        Case Else
            'Not sure what this frame is, stop reading Tag for fear of a read error...
            'If bolAutoReadUnknown = False Then
            'lngResult = MsgBox("There was an unknown Frame " & Chr(34) & Replace(Frame.strID, Chr(0), " ") & Chr(34) & " encountered.  This program cannot identify this frame and would like to stop reading the Tag.  Do you wish to proceed anyway?", vbYesNo + vbDefaultButton2 + vbInformation)
            'If lngResult = vbNo Then Exit Do
            dblFrameSize = DecodeUnsync(Frame.bytSize) - 1
            'End If
    End Select
    ReDim bytFrameData(dblFrameSize)
    Get #intFile, , bytFrameData
    dblBytesRead = dblBytesRead + dblFrameSize + 1
    If (bolCompressed = True) Or (bolEncrypted = True) Then GoTo NextFrame
    With Frame
        
        Select Case .strID
            Case "TIT2"
                'Song Title
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                SongTitle = Right(strData, Len(strData) - 1)
            Case "TPE1"
                'Artist
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Artist = Right(strData, Len(strData) - 1)
            Case "TALB"
                'Album
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Album = Right(strData, Len(strData) - 1)
            Case "TYER"
                'Year
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Year = Right(strData, Len(strData) - 1)
            Case "COMM"
                'Comment
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                strData = Right(strData, Len(strData) - 1)
                strData = Replace(strData, Chr(0), "{::}")
                strCommentInfo = Split(strData, "{::}")
                Comment = strCommentInfo(1)
            Case "TCON"
                'Genre
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Genre = Right(strData, Len(strData) - 1)
            Case "TRCK"
                'Track/Tracks
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                strData = Right(strData, Len(strData) - 1)
                If InStr(strData, "/") <> 0 Then
                    strTrackInfo = Split(strData, "/")
                    If IsNumeric(strTrackInfo(0)) = True Then Track = strTrackInfo(0)
                    If IsNumeric(strTrackInfo(1)) = True Then Tracks = strTrackInfo(1)
                Else
                    If IsNumeric(strData) = True Then Track = CByte(strData)
                End If
            Case "TENC"
                'Encoded By
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                EncodedBy = Right(strData, Len(strData) - 1)
            Case "TCOM"
                'Composer
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Composer = Right(strData, Len(strData) - 1)
            Case "TOPE"
                'Original Artist
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                OriginalArtist = Right(strData, Len(strData) - 1)
            Case "TCOP"
                'Copyright
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Copyright = Right(strData, Len(strData) - 1)
            Case "TIT3"
                'SubTitle
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                SubTitle = Right(strData, Len(strData) - 1)
            Case "TLAN"
                'Language
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                Language = Right(strData, Len(strData) - 1)
            Case "WXXX"
                'URL
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                strData = Right(strData, Len(strData) - 1)
                strData = Replace(strData, Chr(0), "{::}")
                strURLInfo = Split(strData, "{::}")
                If URL = "" Then
                    URL = IIf(strURLInfo(0) = "", strURLInfo(1), strURLInfo(0) & "->" & strURLInfo(1))
                Else
                    URL = URL & "|" & IIf(strURLInfo(0) = "", strURLInfo(1), strURLInfo(0) & "->" & strURLInfo(1))
                End If
            Case "USLT"
                'Lyrics
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                strData = Right(strData, Len(strData) - 1)
                strData = Replace(strData, Chr(0), "{::}")
                strLyricsInfo = Split(strData, "{::}")
                Select Case UBound(strLyricsInfo)
                    Case 1
                        Lyrics = strLyricsInfo(UBound(strLyricsInfo))
                    Case 2
                        Lyrics = strLyricsInfo(UBound(strLyricsInfo) - 1) & vbCrLf & strLyricsInfo(UBound(strLyricsInfo))
                End Select
            Case "SYLT"
                'Sync Lyrics
                SyncLyrics = DecodeSyncLyrics(bytFrameData)
            Case "APIC"
                strDescriptor = ""
                bytTextEncoding = bytFrameData(0)
                ExtractImageInfo bytFrameData, "", bytInternalImageType, strDescriptor
                If GraphicExtended = "" Then
                    GraphicExtended = GetImageType(bytInternalImageType)
                    GraphicTitle = strDescriptor
                Else
                    GraphicExtended = GraphicExtended & "|" & GetImageType(bytInternalImageType)
                    GraphicTitle = GraphicTitle & "|" & strDescriptor
                End If
            Case Else
                'The unknown frame
                strData = GetDataFromBytes(bytFrameData)
                bytTextEncoding = Asc(Left(strData, 1))
                If Unknown = "" Then
                    Unknown = Chr(34) & .strID & Chr(34) & "=" & Right(Replace(strData, Chr(0), " "), Len(strData) - 1)
                Else
                    Unknown = Unknown & "|" & Chr(34) & .strID & Chr(34) & "=" & Right(Replace(strData, Chr(0), " "), Len(strData) - 1)
                End If
        End Select
    End With
    strData = ""
Loop Until (dblTotalTagSize - dblBytesRead) <= 0
Close intFile
TagPresent = True
Exit Sub
ErrHan:
TagPresent = False
Close intFile
End Sub

Private Function GetImageType(ByVal bytType As Byte) As String
Select Case bytType
    Case 1
        GetImageType = "32x32 Icon"
    Case 2
        GetImageType = "Other Icon"
    Case 3
        GetImageType = "Front Cover"
    Case 4
        GetImageType = "Back Cover"
    Case 5
        GetImageType = "Leaflet Page"
    Case 6
        GetImageType = "Media"
    Case 7
        GetImageType = "Lead Performer"
    Case 8
        GetImageType = "Artist"
    Case 9
        GetImageType = "Conductor"
    Case 10
        GetImageType = "Band"
    Case 11
        GetImageType = "Composer"
    Case 12
        GetImageType = "Lyricist"
    Case 13
        GetImageType = "Recording Location"
    Case 14
        GetImageType = "During Recording"
    Case 15
        GetImageType = "During Performance"
    Case 16
        GetImageType = "Movie Capture"
    Case 17
        GetImageType = "A Bright Colored Fish???"
    Case 18
        GetImageType = "Illustration"
    Case 19
        GetImageType = "Band Logo"
    Case 20
        GetImageType = "Publisher Logo"
    Case Else
        GetImageType = "Other"
End Select
End Function

Private Function GetDataFromBytes(ByRef bytBytes() As Byte) As String
Dim lngCounter As Long
Dim strData As String
On Local Error GoTo ErrHan
For lngCounter = 0 To UBound(bytBytes)
    strData = strData & Chr(bytBytes(lngCounter))
Next
GetDataFromBytes = strData
Exit Function
ErrHan:
GetDataFromBytes = ""
End Function

Private Sub DecodeFrameFlags(ByRef bytBytes() As Byte, ByRef bolTagAlterDiscard As Boolean, ByRef bolFileAlterDiscard As Boolean, ByRef bolReadOnly As Boolean, ByRef bolCompressed As Boolean, ByRef bolEncrypted As Boolean, ByRef bolGrouped As Boolean)
Dim strBinary As String
On Error GoTo ErrHan
bolTagAlterDiscard = False
bolFileAlterDiscard = False
bolReadOnly = False
bolCompressed = False
bolEncrypted = False
bolGrouped = False
If bytBytes(0) <> 0 Then
    strBinary = NumberToBinary(bytBytes(0))
    If Mid(strBinary, 1, 1) = "1" Then bolTagAlterDiscard = True
    If Mid(strBinary, 2, 1) = "1" Then bolFileAlterDiscard = True
    If Mid(strBinary, 3, 1) = "1" Then bolReadOnly = True
End If
If bytBytes(1) <> 0 Then
    strBinary = NumberToBinary(bytBytes(1))
    If Mid(strBinary, 1, 1) = "1" Then bolCompressed = True
    If Mid(strBinary, 2, 1) = "1" Then bolEncrypted = True
    If Mid(strBinary, 3, 1) = "1" Then bolGrouped = True
End If
Exit Sub
ErrHan:
End Sub

Private Sub GetFlagInformation(ByVal bytByte As Byte, ByRef bolUnsync As Boolean, ByRef bolExtended As Boolean, ByRef bolExperimental As Boolean)
Dim strBinary As String
On Error Resume Next
strBinary = NumberToBinary(bytByte)
bolUnsync = IIf(Mid(strBinary, 1, 1) = "1", True, False)
bolExtended = IIf(Mid(strBinary, 2, 1) = "1", True, False)
bolExperimental = IIf(Mid(strBinary, 3, 1) = "1", True, False)
End Sub

Public Function ExtractGraphic(ByVal strFileName As String, ByVal lngStartPosition As Long, ByVal lngSize As Long, Optional ByRef strGraphicInfo As String, Optional ByRef strGraphicType As String) As String
Dim intFile As Integer
Dim bytBytes() As Byte
Dim lngBuffer As Long
Dim lngBytesRead As Long
Dim lngBytesLeft As Long
Dim intOutFile As Integer
Dim strOutFile As String
Dim strImageType As String
Dim strDescriptor As String
Dim lngCounter As Long
Dim bytInternalType As Byte
Dim dblOffset As Double
On Error GoTo ErrHan
If FileExists(strFileName) = False Then Exit Function
intFile = FreeFile
lngBuffer = 2097152
lngSize = lngSize - 1
lngBytesLeft = lngSize
lngBytesRead = 0
dblOffset = lngStartPosition + 11
Open strFileName For Binary Access Read Shared As intFile
intOutFile = FreeFile
Open strFileName & ".pic" For Binary Access Write As intOutFile
Do
    If lngBytesLeft >= lngBuffer Then
        ReDim bytBytes(lngBuffer)
    Else
        lngBuffer = lngBytesLeft
        ReDim bytBytes(lngBuffer)
    End If
    Get #intFile, dblOffset + lngBytesRead, bytBytes
    If lngBytesLeft = lngSize Then
        ExtractImageInfo bytBytes, strImageType, bytInternalType, strDescriptor
        Put #intOutFile, , bytBytes
        dblOffset = dblOffset + 1
    Else
        Put #intOutFile, , bytBytes
        dblOffset = dblOffset + 1
    End If
    lngBytesRead = lngBytesRead + lngBuffer
    lngBytesLeft = lngBytesLeft - lngBuffer
    DoEvents
Loop Until lngBytesLeft <= 0
Close intOutFile
Close intFile
Select Case strImageType
    Case "jpeg"
        lngCounter = 0
        Do Until FileExists(strFileName & lngCounter & ".jpg") = False
            lngCounter = lngCounter + 1
        Loop
        Name strFileName & ".pic" As strFileName & lngCounter & ".jpg"
        ExtractGraphic = strFileName & lngCounter & ".jpg"
    Case Else
        lngCounter = 0
        Do Until FileExists(strFileName & lngCounter & "." & strImageType) = False
            lngCounter = lngCounter + 1
        Loop
        Name strFileName & ".pic" As strFileName & lngCounter & "." & strImageType
        ExtractGraphic = strFileName & lngCounter & "." & strImageType
End Select
strGraphicType = strImageType
strGraphicInfo = strDescriptor
Exit Function
ErrHan:
End Function

Private Sub ExtractImageInfo(ByRef bytBytes() As Byte, ByRef strImageType As String, ByRef bytInternalType As Byte, ByRef strDescriptor As String)
Dim lngCounter As Long
Dim intImageFlag As Integer
Dim intDescriptorFlag As Integer
Dim lngBytesRead As Long
Dim bytNewBytes() As Byte
On Error GoTo ErrHan
For lngCounter = 1 To UBound(bytBytes)
    Select Case bytBytes(lngCounter)
        Case 0
            If intImageFlag = 0 Then
                intImageFlag = 1
                lngCounter = lngCounter + 1
                bytInternalType = bytBytes(lngCounter)
            Else
                If intImageFlag = 1 Then
                    If intDescriptorFlag = 0 Then
                        intDescriptorFlag = 1
                    Else
                        Exit For
                    End If
                End If
            End If
        Case Else
            If intImageFlag = 0 Then
                strImageType = strImageType & Chr(bytBytes(lngCounter))
            Else
                If intDescriptorFlag = 0 Then
                    strDescriptor = strDescriptor & Chr(bytBytes(lngCounter))
                Else
                    Exit For
                End If
            End If
    End Select
Next
lngBytesRead = lngCounter
ReDim bytNewBytes(UBound(bytBytes) - lngBytesRead)
For lngCounter = lngBytesRead To UBound(bytBytes)
    bytNewBytes(lngCounter - lngBytesRead) = bytBytes(lngCounter)
Next
ReDim bytBytes(UBound(bytNewBytes))
For lngCounter = 0 To UBound(bytNewBytes)
    bytBytes(lngCounter) = bytNewBytes(lngCounter)
Next
strImageType = Replace(strImageType, "image/", "")
Exit Sub
ErrHan:
strImageType = ""
strDescriptor = ""
End Sub

Private Function LeftShift(ByVal intNumber As Integer, ByVal bytAmount As Byte) As Long
Dim strTemp As String
If (bytAmount = 0) Or (intNumber = 0) Then
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

Private Function DecodeUnsync(ByRef bytBytes() As Byte) As Double
On Error GoTo ErrHan
DecodeUnsync = LeftShift(bytBytes(0), 21) + LeftShift(bytBytes(1), 14) + LeftShift(bytBytes(2), 7) + bytBytes(3)
Exit Function
ErrHan:
DecodeUnsync = 0
End Function

Private Function DecodeSync(ByRef bytBytes() As Byte) As Double
On Error GoTo ErrHan
DecodeSync = BinaryToLong(NumberToBinary(bytBytes(0)) & NumberToBinary(bytBytes(1)) & NumberToBinary(bytBytes(2)) & NumberToBinary(bytBytes(3)))
Exit Function
ErrHan:
DecodeSync = 0
End Function

Private Function FileExists(ByVal strFileName As String) As Boolean
Dim intFile As Integer
On Error GoTo ErrHan
intFile = FreeFile
Open strFileName For Input As intFile
Close
FileExists = True
Exit Function
ErrHan:
FileExists = False
End Function

Private Sub ClearProperties()
SongTitle = ""
Artist = ""
Album = ""
Year = ""
Comment = ""
Track = 0
Tracks = 0
Genre = ""
URL = ""
Composer = ""
OriginalArtist = ""
Copyright = ""
EncodedBy = ""
Lyrics = ""
GraphicStartPos = ""
GraphicSize = ""
GraphicExtended = ""
SubTitle = ""
Language = ""
Unknown = ""
GraphicTitle = ""
End Sub

Private Function DecodeSyncLyrics(ByRef bytBytes() As Byte) As String
Dim bytEncoding As Byte
Dim strLanguage As String
Dim bytTimeStampFormat As Byte
Dim bytContentType As Byte
Dim strDescriptor As String
Dim lngCounter As Long
Dim strTemp As String
Dim strLyrics As String
Dim lngMilliseconds As Long
Dim bytSync() As Byte
On Local Error GoTo ErrHan
ReDim bytSync(0 To 3)
bytEncoding = bytBytes(0)
strLanguage = Chr(bytBytes(1)) & Chr(bytBytes(2)) & Chr(bytBytes(3))
bytTimeStampFormat = bytBytes(4)
bytContentType = bytBytes(5)
'Descriptor
lngCounter = 6
Do Until bytBytes(lngCounter) = 0
    strDescriptor = strDescriptor & Chr(bytBytes(lngCounter))
    lngCounter = lngCounter + 1
Loop
'Sync Lyrics
lngCounter = lngCounter + 1
Do Until lngCounter >= UBound(bytBytes)
    If bytEncoding = 1 Then
        If (bytBytes(lngCounter) = 0) And (bytBytes(lngCounter + 1) = 0) Then
            lngCounter = lngCounter + 1
            bytSync(0) = bytBytes(lngCounter + 1)
            bytSync(1) = bytBytes(lngCounter + 2)
            bytSync(2) = bytBytes(lngCounter + 3)
            bytSync(3) = bytBytes(lngCounter + 4)
            lngMilliseconds = DecodeSync(bytSync)
            lngCounter = lngCounter + 4
            If strLyrics = "" Then
                strLyrics = "[" & MakeTime(lngMilliseconds) & "]" & strTemp
            Else
                strLyrics = strLyrics & "[" & MakeTime(lngMilliseconds) & "]" & strTemp
            End If
            strTemp = ""
        Else
            If bytBytes(lngCounter) = 10 Then
                strTemp = strTemp & vbCrLf
            Else
                strTemp = strTemp & Chr(bytBytes(lngCounter))
            End If
        End If
    Else
        If bytBytes(lngCounter) = 0 Then
            bytSync(0) = bytBytes(lngCounter + 1)
            bytSync(1) = bytBytes(lngCounter + 2)
            bytSync(2) = bytBytes(lngCounter + 3)
            bytSync(3) = bytBytes(lngCounter + 4)
            lngMilliseconds = DecodeSync(bytSync)
            lngCounter = lngCounter + 4
            If strLyrics = "" Then
                strLyrics = "[" & MakeTime(lngMilliseconds) & "]" & strTemp
            Else
                strLyrics = strLyrics & "[" & MakeTime(lngMilliseconds) & "]" & strTemp
            End If
            strTemp = ""
        Else
            If bytBytes(lngCounter) = 10 Then
                strLyrics = strLyrics & vbCrLf
            Else
                strTemp = strTemp & Chr(bytBytes(lngCounter))
            End If
        End If
    End If
    lngCounter = lngCounter + 1
Loop
DecodeSyncLyrics = strLyrics
Exit Function
ErrHan:
MsgBox Err.Description
End Function

Private Function MakeTime(ByVal lngMill As Long) As String
Dim lngSec As Long
Dim lngMin As Long
lngSec = lngMill \ 1000
lngMin = lngSec \ 60
lngSec = lngSec - (lngMin * 60)
MakeTime = Format(lngMin, "00") & ":" & Format(lngSec, "00")
End Function

