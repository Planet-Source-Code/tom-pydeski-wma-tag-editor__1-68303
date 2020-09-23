VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type MS
    M As Integer
    S As Integer
End Type
Private MP3Info As New clsMP3HeadReader
Private WMAInfo As New clsWMAHeadReader

Private Function Get_Extension(strFileName As String) As String
Get_Extension = UCase(Right(strFileName, (Len(strFileName) - InStrRev(strFileName, ".", Len(strFileName)))))
End Function

Function Find_Header_Tags(strFileName As String) As Boolean
On Error GoTo Oops
Set InfoTag = New InfoTag
Dim MP3Time As MS
Dim strDisplay As String
Clear_Properties
Select Case UCase(Get_Extension(strFileName))
    Case "MP3"
        MP3Info.ReadMP3Header strFileName
        'MP3Time = TMS(MP3Info.Length)
        MP3Info.LoadID3Info strFileName
        'strDisplay = "Bitrate: " & MP3Info.Bitrate & vbCrLf & "Length: " & MP3Time.M & ":" & Format(MP3Time.S, "00") & vbCrLf & "Copyright: " & MP3Info.IsCopyright & vbCrLf & "Protected: " & MP3Info.IsProtected & vbCrLf & "Size: " & MP3Info.FileSize & vbCrLf
        'strDisplay = strDisplay & "Title: " & MP3Info.Title & vbCrLf & "Artist: " & MP3Info.Artist & vbCrLf & "Album: " & MP3Info.Album & vbCrLf & "Year: " & MP3Info.Year & vbCrLf & "Comments: " & MP3Info.Comments
        Find_Header_Tags = True
    Case "WMA"
        'MsgBox "WMA Analyzing " & strFileName
        WMAInfo.Get_WMA_Header strFileName
        Find_Header_Tags = True
    Case Else
        MsgBox strFileName & " is not supported by this TagEditor!"
        Find_Header_Tags = False
End Select
GoTo Exit_Find_Header_Tags
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Find_Header_Tags "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Find_Header_Tags"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Find_Header_Tags:
End Function

Private Function TMS(ByVal lngSeconds As Long) As MS
TMS.M = lngSeconds \ 60
lngSeconds = lngSeconds - (TMS.M * 60)
TMS.S = lngSeconds
End Function

Private Function Clear_Properties()
fstrMPEG = ""
fbytLayer = 0
fbolIsProtected = False
fintBitrate = 0
flngFrequency = 0
fbolIsPublic = False
fstrMode = ""
fbolIsCopyright = False
fbolIsLicensed = False
flngLength = 0
fstrArtist = ""
fstrAlbum = ""
fstrTitle = ""
fstrYear = ""
fstrComments = ""
fstrFileSize = ""
fstrHeader = ""
fstrHeaderVersion = ""
fstrHeaderTail = ""
fintLayer = 0
'below added by tp
fintTrackNumber = 0
fstrComposer = ""
fstrPublisher = ""
fstrGenre = ""
End Function
'Property Let Mode(ByVal strMode As String)
'    fstrMode = strMode
'End Property

Property Get Mode() As String
Mode = fstrMode
End Property
'Property Let Header(ByVal strHeader As String)
'    fstrHeader = strHeader
'End Property

Property Get Header() As String
Header = fstrHeader
End Property
'Property Let Layer(ByVal intLayer As Integer)
'    fintLayer = intLayer
'End Property

Property Get Layer() As Integer
Layer = fintLayer
End Property
'Property Let Frequency(ByVal strFrequency As String)
'    flngFrequency = strFrequency
'End Property

Property Get Frequency() As String
Frequency = flngFrequency
End Property
'Property Let FileSize(ByVal strFileSize As String)
'    fstrFileSize = strFileSize
'End Property

Property Get FileSize() As String
FileSize = fstrFileSize
End Property
'Property Let Artist(ByVal strArtist As String)
'    fstrArtist = strArtist
'End Property

Property Get Artist() As String
Artist = fstrArtist
End Property
'Property Let Album(ByVal strAlbum As String)
'    fstrAlbum = strAlbum
'End Property

Property Get Album() As String
Album = fstrAlbum
End Property
'Property Let Title(ByVal strTitle As String)
'    fstrTitle = strTitle
'End Property

Property Get Title() As String
Title = fstrTitle
End Property
'Property Let Year(ByVal strYear As String)
'    fstrYear = strYear
'End Property

Property Get Year() As String
Year = fstrYear
End Property
'Property Let Comments(ByVal strComments As String)
'    fstrComments = strComments
'End Property

Property Get Comments() As String
Comments = fstrComments
End Property
'Property Let IsProtected(ByVal bolIsProtected As Boolean)
'    fbolIsProtected = bolIsProtected
'End Property

Property Get IsProtected() As Boolean
IsProtected = fbolIsProtected
End Property
'Property Let Bitrate(ByVal intBitrate As Long)
'    fintBitrate = intBitrate
'End Property

Property Get Bitrate() As Long
Bitrate = fintBitrate
End Property
'Property Let IsCopyright(ByVal bolIsCopyright As Boolean)
'    fbolIsCopyright = bolIsCopyright
'End Property

Property Get IsCopyright() As Boolean
IsCopyright = fbolIsCopyright
End Property
'Property Let Length(ByVal lngLength As Long)
'    flngLength = lngLength
'End Property

Property Get Length() As Long
Length = flngLength
End Property
'Property Let IsLicensed(ByVal bolIsLicensed As Boolean)
'    fbolIsLicensed = bolIsLicensed
'End Property

Property Get IsLicensed() As Boolean
IsLicensed = fbolIsLicensed
End Property

'below added by tp
'Property Let Composer(ByVal strComposer As string)
'    fstrComposer = strComposer
'End Property

Property Get Composer() As String
Composer = fstrComposer
End Property

'Property Let Publisher(ByVal strPublisher As string)
'    fstrPublisher = strPublisher
'End Property

Property Get Publisher() As String
Publisher = fstrPublisher
End Property

'Property Let Genre(ByVal strGenre As string)
'    fstrGenre = strGenre
'End Property

Property Get Genre() As String
Genre = fstrGenre
End Property

'Property Let TrackNumber(ByVal bytTrackNumber As Integer)
'    fbyteTrackNumber = intTrackNumber
'End Property

Property Get TrackNumber() As Integer
TrackNumber = fbyteTrackNumber
End Property
