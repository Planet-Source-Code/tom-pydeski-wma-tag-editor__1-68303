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
Attribute VB_Exposed = False
'modified from Somenon's original submission
Private Type MS
    M As Integer
    S As Integer
End Type
Private WMAInfo As New clsWMAHeadReader

Private Function Get_Extension(strFileName As String) As String
Get_Extension = UCase(Right(strFileName, (Len(strFileName) - InStrRev(strFileName, ".", Len(strFileName)))))
End Function

Function Find_Header_Tags(strFileName As String) As Boolean
On Error GoTo Oops
'initialize our tag class
Set InfoTag = New InfoTag
Dim MP3Time As MS
Dim strDisplay As String
Clear_Properties
Select Case UCase(Get_Extension(strFileName))
    Case "WMA"
        'MsgBox "WMA Analyzing " & strFileName
        'WMAInfo.Get_WMA_Header strFileName
        WMAInfo.ReadWMAHeader strFileName
        Find_Header_Tags = True
    Case Else
        MsgBox strFileName & " is not supported by this TagEditor!"
        Find_Header_Tags = False
        WMAInfo.ReadWMAHeader strFileName
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

Sub Save_Header_Tags(strFileName As String)
WMAInfo.SaveWMAHeader strFileName
End Sub

Private Function TMS(ByVal lngSeconds As Long) As MS
TMS.M = lngSeconds \ 60
lngSeconds = lngSeconds - (TMS.M * 60)
TMS.S = lngSeconds
End Function

Private Function Clear_Properties()
fstrMPEG = ""
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
fstrFullHeader = ""
fstrHeaderVersion = ""
fstrHeaderTail = ""
fintLayer = 0
'below added by tp
fstrDuration = ""
fintTrackNumber = 0
fstrComposer = ""
fstrPublisher = ""
fstrGenre = ""
fstrAlbumArtist = ""
fbyteTrackNumber = 0
fstrDescription = ""
fstrRating = ""
fstrProviderRating = ""
End Function

Public Property Let Mode(ByVal strMode As String)
fstrMode = strMode
End Property

Public Property Get Mode() As String
Mode = fstrMode
End Property

Public Property Let Header(ByVal strHeader As String)
fstrHeader = strHeader
End Property

Public Property Get Header() As String
Header = fstrHeader
End Property

Public Property Let Layer(ByVal intLayer As Integer)
fintLayer = intLayer
End Property

Public Property Get Layer() As Integer
Layer = fintLayer
End Property

Public Property Let Frequency(ByVal strFrequency As String)
flngFrequency = strFrequency
End Property

Public Property Get Frequency() As String
Frequency = flngFrequency
End Property

Public Property Let FileSize(ByVal strFileSize As String)
fstrFileSize = strFileSize
End Property

Public Property Get FileSize() As String
FileSize = fstrFileSize
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

Public Property Let Title(ByVal strTitle As String)
fstrTitle = strTitle
End Property

Public Property Get Title() As String
Title = fstrTitle
End Property

Public Property Let Year(ByVal strYear As String)
fstrYear = strYear
End Property

Public Property Get Year() As String
Year = fstrYear
End Property

Public Property Let Comments(ByVal strComments As String)
fstrComments = strComments
End Property

Public Property Get Comments() As String
Comments = fstrComments
End Property

Public Property Let IsProtected(ByVal bolIsProtected As Boolean)
fbolIsProtected = bolIsProtected
End Property

Public Property Get IsProtected() As Boolean
IsProtected = fbolIsProtected
End Property

Public Property Let Bitrate(ByVal intBitrate As Long)
fintBitrate = intBitrate
End Property

Public Property Get Bitrate() As Long
Bitrate = fintBitrate
End Property

Public Property Let IsCopyright(ByVal bolIsCopyright As Boolean)
fbolIsCopyright = bolIsCopyright
End Property

Public Property Get IsCopyright() As Boolean
IsCopyright = fbolIsCopyright
End Property

Public Property Let Length(ByVal lngLength As Long)
flngLength = lngLength
End Property

Public Property Get Length() As Long
Length = flngLength
End Property

Public Property Let IsLicensed(ByVal bolIsLicensed As Boolean)
fbolIsLicensed = bolIsLicensed
End Property

Public Property Get IsLicensed() As Boolean
IsLicensed = fbolIsLicensed
End Property
'============================================================================
'below added by tp

Public Property Let Composer(ByVal strComposer As String)
fstrComposer = strComposer
End Property

Public Property Get Composer() As String
Composer = fstrComposer
End Property

Public Property Let Publisher(ByVal strPublisher As String)
fstrPublisher = strPublisher
End Property

Public Property Get Publisher() As String
Publisher = fstrPublisher
End Property

Public Property Let Genre(ByVal strGenre As String)
fstrGenre = strGenre
End Property

Public Property Get Genre() As String
Genre = fstrGenre
End Property

Public Property Let TrackNumber(ByVal bytTrackNumber As Long)
fbyteTrackNumber = bytTrackNumber
End Property

Public Property Get TrackNumber() As Long
TrackNumber = fbyteTrackNumber
End Property

Public Property Let FullHeader(ByVal strFullHeader As String)
fstrFullHeader = strFullHeader
End Property

Public Property Get FullHeader() As String
FullHeader = fstrFullHeader
End Property

Public Property Let Duration(ByVal strDuration As String)
fstrDuration = strDuration
End Property

Public Property Get Duration() As String
Duration = fstrDuration
End Property

Public Property Let Info(ByVal strInfo As String)
fstrInfo = strInfo
End Property

Public Property Get Info() As String
Info = fstrInfo
End Property

Public Property Let Description(ByVal strDescription As String)
fstrDescription = strDescription
End Property

Public Property Get Description() As String
Description = fstrDescription
End Property

Public Property Let Rating(ByVal strRating As String)
fstrRating = strRating
End Property

Public Property Get Rating() As String
Rating = fstrRating
End Property

Public Property Let ProviderRating(ByVal strProviderRating As String)
fstrProviderRating = strProviderRating
End Property

Public Property Get ProviderRating() As String
ProviderRating = fstrProviderRating
End Property

Public Property Let Copyright(ByVal strCopyright As String)
fstrCopyright = strCopyright
End Property

Public Property Get Copyright() As String
Copyright = fstrCopyright
End Property

Public Property Let AlbumArtist(ByVal strAlbumArtist As String)
fstrAlbumArtist = strAlbumArtist
End Property

Public Property Get AlbumArtist() As String
AlbumArtist = fstrAlbumArtist
End Property
