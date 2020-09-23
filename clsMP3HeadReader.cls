VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3HeadReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Graphics
    strFileName As String
    bytType As Byte
    strTitle As String
    dblStartPosition As Double
    dblLength As Double
    bolExtracted As Boolean
End Type
Private Type Lyric
    strLine As String
    intMinute As Integer
    intSecond As Integer
End Type
Private Lyrics() As Lyric
Private fbolStop As Boolean
Private fbolStopBatch As Boolean
Private fGraphics() As Graphics

Public Sub ReadMP3Header(ByVal strFileName As String)
Dim intFile As Integer
Dim bytByte As Byte
Dim bytArray() As Byte
Dim strBits As String
Dim lngCounter As Long
Dim lngTagSize As Long
On Error GoTo ErrHan
If FileExists(strFileName) = False Then Exit Sub

fstrFileSize = FileLen(strFileName)
lngTagSize = IDTagSize(strFileName)
intFile = FreeFile
Open strFileName For Binary Access Read Shared As intFile
Do Until lngCounter >= lngTagSize
    Get #intFile, , bytByte
    lngCounter = lngCounter + 1
Loop
Do Until EOF(intFile) Or (bytByte = 255)
    Get #intFile, , bytByte
Loop
If bytByte = 255 Then
    ReDim bytArray(1 To 4)
    bytArray(1) = 255
    Get #intFile, , bytArray(2)
    Get #intFile, , bytArray(3)
    Get #intFile, , bytArray(4)
End If
Close
strBits = NumberToBinary(bytArray(1))
strBits = strBits & NumberToBinary(bytArray(2))
strBits = strBits & NumberToBinary(bytArray(3))
strBits = strBits & NumberToBinary(bytArray(4))
ReadMP3Bits strBits, strFileName
Exit Sub
ErrHan:
End Sub

Private Sub ReadMP3Bits(ByVal strBits As String, ByVal strFileName As String)
Dim strSync As String
Dim strID As String
Dim strLayer As String
Dim strProtected As String
Dim strBitrate As String
Dim strFrequency As String
Dim strPad As String
Dim strPrivate As String
Dim strMode As String
Dim strModeExtension As String
Dim strCopyright As String
Dim MPEG As String
If Len(strBits) <> 32 Then Exit Sub
On Error GoTo ErrHan
strSync = Mid(strBits, 1, 12)
strID = Mid(strBits, 13, 1)
strLayer = Mid(strBits, 14, 2)
strProtected = Mid(strBits, 16, 1)
strBitrate = Mid(strBits, 17, 4)
strFrequency = Mid(strBits, 21, 2)
strPad = Mid(strBits, 23, 1)
strPrivate = Mid(strBits, 24, 1)
strMode = Mid(strBits, 25, 2)
strModeExtension = Mid(strBits, 27, 2)
strCopyright = Mid(strBits, 29, 1)

Select Case strLayer
    Case "01"
        fbytLayer = 3
    Case "10"
        fbytLayer = 2
    Case "11"
        fbytLayer = 1
    Case Else
        fbytLayer = 0
End Select

If strID = "1" Then
    MPEG = "MPEG-1"
    Select Case strFrequency
        Case "00"
            flngFrequency = 44100
        Case "01"
            flngFrequency = 48000
        Case "10"
            flngFrequency = 32000
        Case Else
            flngFrequency = 0
    End Select
    Select Case strBitrate
        Case "0001"
            fintBitrate = 32
        Case "0010"
            If fbytLayer = 1 Then
                fintBitrate = 64
            ElseIf fbytLayer = 2 Then
                fintBitrate = 48
            ElseIf fbytLayer = 3 Then
                fintBitrate = 40
            End If
        Case "0011"
            If fbytLayer = 1 Then
                fintBitrate = 96
            ElseIf fbytLayer = 2 Then
                fintBitrate = 56
            ElseIf fbytLayer = 3 Then
                fintBitrate = 48
            End If
        Case "0100"
            If fbytLayer = 1 Then
                fintBitrate = 128
            ElseIf fbytLayer = 2 Then
                fintBitrate = 64
            ElseIf fbytLayer = 3 Then
                fintBitrate = 56
            End If
        Case "0101"
            If fbytLayer = 1 Then
                fintBitrate = 160
            ElseIf fbytLayer = 2 Then
                fintBitrate = 80
            ElseIf fbytLayer = 3 Then
                fintBitrate = 64
            End If
        Case "0110"
            If fbytLayer = 1 Then
                fintBitrate = 192
            ElseIf fbytLayer = 2 Then
                fintBitrate = 96
            ElseIf fbytLayer = 3 Then
                fintBitrate = 80
            End If
        Case "0111"
            If fbytLayer = 1 Then
                fintBitrate = 224
            ElseIf fbytLayer = 2 Then
                fintBitrate = 112
            ElseIf fbytLayer = 3 Then
                fintBitrate = 96
            End If
        Case "1000"
            If fbytLayer = 1 Then
                fintBitrate = 256
            ElseIf fbytLayer = 2 Then
                fintBitrate = 128
            ElseIf fbytLayer = 3 Then
                fintBitrate = 112
            End If
        Case "1001"
            If fbytLayer = 1 Then
                fintBitrate = 288
            ElseIf fbytLayer = 2 Then
                fintBitrate = 160
            ElseIf fbytLayer = 3 Then
                fintBitrate = 128
            End If
        Case "1010"
            If fbytLayer = 1 Then
                fintBitrate = 320
            ElseIf fbytLayer = 2 Then
                fintBitrate = 192
            ElseIf fbytLayer = 3 Then
                fintBitrate = 160
            End If
        Case "1011"
            If fbytLayer = 1 Then
                fintBitrate = 352
            ElseIf fbytLayer = 2 Then
                fintBitrate = 224
            ElseIf fbytLayer = 3 Then
                fintBitrate = 192
            End If
        Case "1100"
            If fbytLayer = 1 Then
                fintBitrate = 384
            ElseIf fbytLayer = 2 Then
                fintBitrate = 256
            ElseIf fbytLayer = 3 Then
                fintBitrate = 224
            End If
        Case "1101"
            If fbytLayer = 1 Then
                fintBitrate = 416
            ElseIf fbytLayer = 2 Then
                fintBitrate = 320
            ElseIf fbytLayer = 3 Then
                fintBitrate = 256
            End If
        Case "1110"
            If fbytLayer = 1 Then
                fintBitrate = 448
            ElseIf fbytLayer = 2 Then
                fintBitrate = 384
            ElseIf fbytLayer = 3 Then
                fintBitrate = 320
            End If
        Case Else
            fintBitrate = 0
    End Select
Else
    MPEG = "MPEG-2"
    Select Case strFrequency
        Case "00"
            flngFrequency = 22050
        Case "01"
            flngFrequency = 24000
        Case "10"
            flngFrequency = 16000
        Case Else
            flngFrequency = 0
    End Select
    Select Case fintBitrate
        Case "0001"
            If fbytLayer = 1 Then
                fintBitrate = 32
            ElseIf fbytLayer = 2 Then
                fintBitrate = 32
            ElseIf fbytLayer = 3 Then
                fintBitrate = 8
            End If
        Case "0010"
            If fbytLayer = 1 Then
                fintBitrate = 64
            ElseIf fbytLayer = 2 Then
                fintBitrate = 48
            ElseIf fbytLayer = 3 Then
                fintBitrate = 16
            End If
        Case "0011"
            If fbytLayer = 1 Then
                fintBitrate = 96
            ElseIf fbytLayer = 2 Then
                fintBitrate = 56
            ElseIf fbytLayer = 3 Then
                fintBitrate = 24
            End If
        Case "0100"
            If fbytLayer = 1 Then
                fintBitrate = 128
            ElseIf fbytLayer = 2 Then
                fintBitrate = 64
            ElseIf fbytLayer = 3 Then
                fintBitrate = 32
            End If
        Case "0101"
            If fbytLayer = 1 Then
                fintBitrate = 160
            ElseIf fbytLayer = 2 Then
                fintBitrate = 80
            ElseIf fbytLayer = 3 Then
                fintBitrate = 64
            End If
        Case "0110"
            If fbytLayer = 1 Then
                fintBitrate = 192
            ElseIf fbytLayer = 2 Then
                fintBitrate = 96
            ElseIf fbytLayer = 3 Then
                fintBitrate = 80
            End If
        Case "0111"
            If fbytLayer = 1 Then
                fintBitrate = 224
            ElseIf fbytLayer = 2 Then
                fintBitrate = 112
            ElseIf fbytLayer = 3 Then
                fintBitrate = 56
            End If
        Case "1000"
            If fbytLayer = 1 Then
                fintBitrate = 256
            ElseIf fbytLayer = 2 Then
                fintBitrate = 128
            ElseIf fbytLayer = 3 Then
                fintBitrate = 64
            End If
        Case "1001"
            If fbytLayer = 1 Then
                fintBitrate = 288
            ElseIf fbytLayer = 2 Then
                fintBitrate = 160
            ElseIf fbytLayer = 3 Then
                fintBitrate = 128
            End If
        Case "1010"
            If fbytLayer = 1 Then
                fintBitrate = 320
            ElseIf fbytLayer = 2 Then
                fintBitrate = 192
            ElseIf fbytLayer = 3 Then
                fintBitrate = 160
            End If
        Case "1011"
            If fbytLayer = 1 Then
                fintBitrate = 352
            ElseIf fbytLayer = 2 Then
                fintBitrate = 224
            ElseIf fbytLayer = 3 Then
                fintBitrate = 112
            End If
        Case "1100"
            If fbytLayer = 1 Then
                fintBitrate = 384
            ElseIf fbytLayer = 2 Then
                fintBitrate = 256
            ElseIf fbytLayer = 3 Then
                fintBitrate = 128
            End If
        Case "1101"
            If fbytLayer = 1 Then
                fintBitrate = 416
            ElseIf fbytLayer = 2 Then
                fintBitrate = 320
            ElseIf fbytLayer = 3 Then
                fintBitrate = 256
            End If
        Case "1110"
            If fbytLayer = 1 Then
                fintBitrate = 448
            ElseIf fbytLayer = 2 Then
                fintBitrate = 384
            ElseIf fbytLayer = 3 Then
                fintBitrate = 320
            End If
        Case Else
            fintBitrate = 0
    End Select
End If

If strProtected = "1" Then
    fbolIsProtected = True
Else
    fbolIsProtected = False
End If

Select Case strMode
    Case "00"
        fstrMode = "Stereo"
    Case "01"
        fstrMode = "Joint Stereo"
    Case "10"
        fstrMode = "Dual Channel"
    Case "11"
        fstrMode = "Mono"
    Case Else
        fstrMode = ""
End Select
If strCopyright = "1" Then
    fbolIsCopyright = True
Else
    fbolIsCopyright = False
End If
If fintBitrate <> 0 Then
    flngLength = ((FileLen(strFileName) - IDTagSize(strFileName)) * 8) / (fintBitrate * 1000)
Else
    flngLength = 0
End If
Exit Sub
ErrHan:
End Sub

Private Sub ClearProperties()
fbytLayer = 0
fbolIsProtected = False
fintBitrate = 0
flngFrequency = 0
fstrMode = ""
fbolIsCopyright = False
End Sub

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

Private Function IDTagSize(ByVal strFileName As String) As Long
Dim intFile As Integer
Dim strID3 As String * 3
Dim bytByte As Byte
Dim bytSize() As Byte
Dim lngSize As Long
On Error GoTo ErrHan
intFile = FreeFile
Open strFileName For Binary Access Read Shared As intFile
Get #intFile, , strID3
If strID3 = "ID3" Then
    ReDim bytSize(1 To 4)
    Get #intFile, , bytByte
    Get #intFile, , bytByte
    Get #intFile, , bytByte
    Get #intFile, , bytSize
    lngSize = LeftShift(bytSize(1), 21) + LeftShift(bytSize(2), 14) + LeftShift(bytSize(3), 7) + bytSize(4) + 10
End If
Close
IDTagSize = lngSize
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

Function LoadID3Info(ByVal strFile As String)
Dim ID3V11 As New clsID3V111Reader
Dim ID3V230 As New clsID3V230Reader
Dim ID3V230X As New clsID3V230Writer
Dim strURl() As String
Dim strGraphicType() As String
Dim strGraphicLength() As String
Dim strGraphicStartPos() As String
Dim strGraphicTitle() As String
Dim lngCounter As Long
Dim strFileName As String
Dim strFilePart As String
Dim intVersion As Integer
On Local Error GoTo ErrHan

ID3V11.ReadID3V111Tag strFile
'lvwMP3.SelectedItem.SubItems(5) = ID3V11.TagPresent
ID3V230.ReadID3V230Tag strFile
'lvwMP3.SelectedItem.SubItems(6) = ID3V230.TagPresent
If ID3V11.TagPresent = True Then
    intVersion = 1
End If
If ID3V230.TagPresent = True Then
    intVersion = 2
End If

Select Case intVersion
    Case 1
        With ID3V11
            If .TagPresent = True Then
                fstrArtist = .Artist
                fstrAlbum = .Album
                fstrYear = .Year
                fstrComments = .Comment
                fstrTitle = .SongTitle
                'txtArtist.Text = .Artist
                'txtAlbum.Text = .Album
                'txtSong.Text = .SongTitle
                'txtTrack.Text = .Track
                'txtYear.Text = .Year
                'txtComment.Text = .Comment
                'cboGenre.Text = .Genre
            End If
        End With
    Case Else
        With ID3V230
            If .TagPresent = True Then
                fstrArtist = .Artist
                fstrAlbum = .Album
                fstrYear = .Year
                fstrComments = .Comment
                fstrTitle = .SongTitle
                'txtArtist.Text = .Artist
                'txtAlbum.Text = .Album
                'txtSong.Text = .SongTitle
                'txtTrack.Text = .Track
                'txtYear.Text = .Year
                'txtComment.Text = .Comment
                'cboGenre.Text = .Genre
                'txtOriginalArtist.Text = .OriginalArtist
                'txtSubtitle.Text = .SubTitle
                'txtComposer.Text = .Composer
                'txtCopyright.Text = .Copyright
                'txtEncodedBy.Text = .EncodedBy
                'txtLanguage.Text = .Language
            End If
        End With
End Select
Exit Function
ErrHan:
End Function
