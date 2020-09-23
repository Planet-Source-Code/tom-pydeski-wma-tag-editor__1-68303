VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3V111Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type ID3Info
    strHeader As String * 3
    strSongTitle As String * 30
    strArtist As String * 30
    strAlbum As String * 30
    strYear As String * 4
    strComment As String * 28
    bytUnused As Byte
    bytTrack As Byte
    bytGenre As Byte
End Type
Private fstrGenres() As Variant
Private fbolTagPresent As Boolean
Private fstrSongTitle As String
Private fstrArtist As String
Private fstrAlbum As String
Private fstrYear As String
Private fstrComment As String
Private fbytTrack As Byte
Private fstrGenre As String

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

Public Property Let Genre(ByVal strGenre As String)
fstrGenre = strGenre
End Property

Public Property Get Genre() As String
Genre = fstrGenre
End Property

Public Sub ReadID3V111Tag(ByVal strFileName As String, Optional ByVal bolDetermineTagPresent As Boolean = False)
Dim intFile As Integer
Dim ID3 As ID3Info
On Error GoTo ErrHan
If FileExists(strFileName) = False Then Exit Sub
intFile = FreeFile
Open strFileName For Binary Access Read Shared As intFile
Get #intFile, LOF(intFile) - 127, ID3
Close intFile
ClearProperties
With ID3
    If .strHeader = "TAG" Then
        If bolDetermineTagPresent = True Then
            TagPresent = True
            Exit Sub
        End If
        SongTitle = StripNulls(.strSongTitle)
        Artist = StripNulls(.strArtist)
        Album = StripNulls(.strAlbum)
        Year = StripNulls(.strYear)
        Comment = StripNulls(.strComment)
        Track = .bytTrack
        Genre = GetGenre(.bytGenre)
        TagPresent = True
    End If
End With
Exit Sub
ErrHan:
TagPresent = False
End Sub

Private Function StripNulls(ByVal strString As String) As String
On Error GoTo ErrHan
StripNulls = Left(strString, InStr(strString, Chr(0)) - 1)
Exit Function
ErrHan:
StripNulls = strString
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

Private Function GetGenre(ByVal bytByte As Byte) As String
On Error GoTo ErrHan
GetGenre = fstrGenres(bytByte)
Exit Function
ErrHan:
GetGenre = ""
End Function

Private Sub Class_Initialize()
fstrGenres = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", "Hip-Hop", _
"Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", "Rock", _
"Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", "Soundtrack", _
"Euro-Techno", "Ambient", "Trip-Hop", "Vocal", "Jazz+Funk", "Fusion", "Trance", _
"Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", "Noise", _
"AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
"Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial", "Electronic", _
"Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", _
"Top 40", "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", _
"New Wave", "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", _
"Acid Punk", "Acid Jazz", "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", _
"Folk", "Folk/Rock", "National folk", "Swing", "Fast-fusion", "Bebob", "Latin", _
"Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", "Progressive Rock", _
"Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", _
"Easy Listening", "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", _
"Sonata", "Symphony", "Booty Bass", "Primus", "Porn Groove", "Satire", "Slow Jam", _
"Club", "Tango", "Samba", "Folklore", "Ballad", "Powder Ballad", "Rhythmic Soul", _
"Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Capella", "Euro-House", "Dance Hall", _
"Goa", "Drum & Bass", "Club House", "Hardcore", "Terror", "Indie", "BritPop", _
"NegerPunk", "Polsk Punk", "Beat", "Christian Gangsta", "Heavy Metal", "Black Metal", _
"Crossover", "Contemporary C", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", _
"Anime", "JPop", "SynthPop")
End Sub

Private Sub ClearProperties()
SongTitle = ""
Artist = ""
Album = ""
Year = ""
Comment = ""
Track = 0
Genre = ""
End Sub

