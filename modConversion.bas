Attribute VB_Name = "Conversions"
Option Explicit
'submitted by Tom Pydeski
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Type DlngSize
    Low As Long
    High As Long
End Type
Type GuidStruct
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Global strGUID As String * 36
Private Declare Function StringFromGUID2 Lib "OLE32.dll" (ByRef rGUID As Any, ByVal lpsz As String, ByVal cchMax As Long) As Long

Function AddZero(StrIn As String, intDigits As Integer) As String
'we can actually do this by
'AddZero = Format(StrIn$, String(intDigits, "0"))
'but it takes over 3X longer!!!!  format is a hog!
'
StrIn = Trim$(StrIn)
If Len(StrIn) >= intDigits Then
    AddZero = StrIn
    Exit Function
End If
AddZero = String(intDigits - Len(StrIn), "0") & StrIn
End Function

Function AddSpace(StrIn As String, intDigits As Integer) As String
'we can actually do this by
'AddZero = Format(StrIn$, String(intDigits, "0"))
'but it takes over 3X longer!!!!  format is a hog!
'
StrIn = Trim$(StrIn)
If Len(StrIn) >= intDigits Then
    AddSpace = StrIn
    Exit Function
End If
AddSpace = Space(intDigits - Len(StrIn)) & StrIn
End Function

Function LoDWORD(ByVal curQWORD As Currency) As Long
'By Sebastian Mares, sebastian@maresweb.net, 20040527
Dim lngLoDWORD As Long
curQWORD = curQWORD / 10000
CopyMemory lngLoDWORD, curQWORD, 4
LoDWORD = lngLoDWORD
End Function

Function HiDWORD(ByVal curQWORD As Currency) As Long
'By Sebastian Mares, sebastian@maresweb.net, 20040527
Dim lngHiDWORD As Long
curQWORD = curQWORD / 10000
CopyMemory lngHiDWORD, ByVal VarPtr(curQWORD) + 4, 4
HiDWORD = lngHiDWORD
End Function

Function QWORD(ByRef DWord As DlngSize) As Variant
' Function QWORD(ByVal lngHiDWORD As Long, ByVal lngLoDWORD As Long) As Variant
On Error GoTo Oops
Debug.Print DWord.Low, DWord.High
'By Sebastian Mares, sebastian@maresweb.net, 20040527
Dim curQWORD As Currency
CopyMemory curQWORD, DWord.Low, 4
CopyMemory ByVal VarPtr(curQWORD) + 4, DWord.High, 4
QWORD = CDbl(curQWORD) * 10000
GoTo Exit_QWORD
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine QWORD "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in QWORD"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_QWORD:
End Function

Function ExtractDate(Intervals As Variant) As Date
'The value is given as the number of 100-nanosecond intervals since January 1, 1601,
'according to Coordinated Universal Time (Greenwich Mean Time).
Dim varSeconds As Variant
Dim varDays As Variant
varSeconds = CVar(Intervals * ((10 ^ -7)))
varDays = varSeconds / 86400
'Debug.Print varSeconds, varDays
ExtractDate = DateAdd("d", varDays, "1/1/1601")
'Debug.Print "ExtractDate = "; ExtractDate
End Function

Function GetFormatTag(FormatTag As Integer) As String
'retrieve the file format coding based on the format tag
Select Case FormatTag
    Case Is = 0 '(0000H)
        GetFormatTag = "Microsoft Unknown Wave Format"
    Case Is = 1 '(0001H)
        GetFormatTag = "Pulse Code Modulation (PCM)"
    Case Is = 2 '(0002H)
        GetFormatTag = "Microsoft ADPCM"
    Case Is = 3 '(0003H)
        GetFormatTag = "IEEE Float"
    Case Is = 4 '(0004H)
        GetFormatTag = "Compaq Computer VSELP"
    Case Is = 5 '(0005H)
        GetFormatTag = "IBM CVSD"
    Case Is = 6 '(0006H)
        GetFormatTag = "Microsoft A-Law"
    Case Is = 7 '(0007H)
        GetFormatTag = "Microsoft mu-Law"
    Case Is = 8 '(0008H)
        GetFormatTag = "Microsoft DTS"
    Case Is = 16 '(0010H)
        GetFormatTag = "OKI ADPCM"
    Case Is = 17 '(0011H)
        GetFormatTag = "Intel DVI/IMA ADPCM"
    Case Is = 18 '(0012H)
        GetFormatTag = "Videologic MediaSpace ADPCM"
    Case Is = 19 '(0013H)
        GetFormatTag = "Sierra Semiconductor ADPCM"
    Case Is = 20 '(0014H)
        GetFormatTag = "Antex Electronics G.723 ADPCM"
    Case Is = 21 '(0015H)
        GetFormatTag = "DSP Solutions DigiSTD"
    Case Is = 22 '(0016H)
        GetFormatTag = "DSP Solutions DigiFIX"
    Case Is = 23 '(0017H)
        GetFormatTag = "Dialogic OKI ADPCM"
    Case Is = 24 '(0018H)
        GetFormatTag = "MediaVision ADPCM"
    Case Is = 25 '(0019H)
        GetFormatTag = "Hewlett-Packard CU"
    Case Is = 32 '(0020H)
        GetFormatTag = "Yamaha ADPCM"
    Case Is = 33 '(0021H)
        GetFormatTag = "Speech Compression Sonarc"
    Case Is = 34 '(0022H)
        GetFormatTag = "DSP Group TrueSpeech"
    Case Is = 35 '(0023H)
        GetFormatTag = "Echo Speech EchoSC1"
    Case Is = 36 '(0024H)
        GetFormatTag = "Audiofile AF36"
    Case Is = 37 '(0025H)
        GetFormatTag = "Audio Processing Technology APTX"
    Case Is = 38 '(0026H)
        GetFormatTag = "AudioFile AF10"
    Case Is = 39 '(0027H)
        GetFormatTag = "Prosody 1612"
    Case Is = 40 '(0028H)
        GetFormatTag = "LRC"
    Case Is = 48 '(0030H)
        GetFormatTag = "Dolby AC2"
    Case Is = 49 '(0031H)
        GetFormatTag = "Microsoft GSM 6.10"
    Case Is = 50 '(0032H)
        GetFormatTag = "MSNAudio"
    Case Is = 51 '(0033H)
        GetFormatTag = "Antex Electronics ADPCME"
    Case Is = 52 '(0034H)
        GetFormatTag = "Control Resources VQLPC"
    Case Is = 53 '(0035H)
        GetFormatTag = "DSP Solutions DigiREAL"
    Case Is = 54 '(0036H)
        GetFormatTag = "DSP Solutions DigiADPCM"
    Case Is = 55 '(0037H)
        GetFormatTag = "Control Resources CR10"
    Case Is = 56 '(0038H)
        GetFormatTag = "Natural MicroSystems VBXADPCM"
    Case Is = 57 '(0039H)
        GetFormatTag = "Crystal Semiconductor IMA ADPCM"
    Case Is = 58 '(003AH)
        GetFormatTag = "EchoSC3"
    Case Is = 59 '(003BH)
        GetFormatTag = "Rockwell ADPCM"
    Case Is = 60 '(003CH)
        GetFormatTag = "Rockwell Digit LK"
    Case Is = 61 '(003DH)
        GetFormatTag = "Xebec"
    Case Is = 64 '(0040H)
        GetFormatTag = "Antex Electronics G.721 ADPCM"
    Case Is = 65 '(0041H)
        GetFormatTag = "G.728 CELP"
    Case Is = 66 '(0042H)
        GetFormatTag = "MSG723"
    Case Is = 80 '(0050H)
        GetFormatTag = "MPEG Layer-2 or Layer-1"
    Case Is = 82 '(0052H)
        GetFormatTag = "RT24"
    Case Is = 83 '(0053H)
        GetFormatTag = "PAC"
    Case Is = 85 '(0055H)
        GetFormatTag = "MPEG Layer-3"
    Case Is = 89 '(0059H)
        GetFormatTag = "Lucent G.723"
    Case Is = 96 '(0060H)
        GetFormatTag = "Cirrus"
    Case Is = 97 '(0061H)
        GetFormatTag = "ESPCM"
    Case Is = 98 '(0062H)
        GetFormatTag = "Voxware"
    Case Is = 99 '(0063H)
        GetFormatTag = "Canopus Atrac"
    Case Is = 100 '(0064H)
        GetFormatTag = "G.726 ADPCM"
    Case Is = 101 '(0065H)
        GetFormatTag = "G.722 ADPCM"
    Case Is = 102 '(0066H)
        GetFormatTag = "DSAT"
    Case Is = 103 '(0067H)
        GetFormatTag = "DSAT Display"
    Case Is = 105 '(0069H)
        GetFormatTag = "Voxware Byte Aligned"
    Case Is = 112 '(0070H)
        GetFormatTag = "Voxware AC8"
    Case Is = 113 '(0071H)
        GetFormatTag = "Voxware AC10"
    Case Is = 114 '(0072H)
        GetFormatTag = "Voxware AC16"
    Case Is = 115 '(0073H)
        GetFormatTag = "Voxware AC20"
    Case Is = 116 '(0074H)
        GetFormatTag = "Voxware MetaVoice"
    Case Is = 117 '(0075H)
        GetFormatTag = "Voxware MetaSound"
    Case Is = 118 '(0076H)
        GetFormatTag = "Voxware RT29HW"
    Case Is = 119 '(0077H)
        GetFormatTag = "Voxware VR12"
    Case Is = 120 '(0078H)
        GetFormatTag = "Voxware VR18"
    Case Is = 121 '(0079H)
        GetFormatTag = "Voxware TQ40"
    Case Is = 128 '(0080H)
        GetFormatTag = "Softsound"
    Case Is = 129 '(0081H)
        GetFormatTag = "Voxware TQ60"
    Case Is = 130 '(0082H)
        GetFormatTag = "MSRT24"
    Case Is = 131 '(0083H)
        GetFormatTag = "G.729A"
    Case Is = 132 '(0084H)
        GetFormatTag = "MVI MV12"
    Case Is = 133 '(0085H)
        GetFormatTag = "DF G.726"
    Case Is = 134 '(0086H)
        GetFormatTag = "DF GSM610"
    Case Is = 136 '(0088H)
        GetFormatTag = "ISIAudio"
    Case Is = 137 '(0089H)
        GetFormatTag = "Onlive"
    Case Is = 145 '(0091H)
        GetFormatTag = "SBC24"
    Case Is = 146 '(0092H)
        GetFormatTag = "Dolby AC3 SPDIF"
    Case Is = 147 '(0093H)
        GetFormatTag = "MediaSonic G.723"
    Case Is = 148 '(0094H)
        GetFormatTag = "Aculab PLC    Prosody 8kbps"
    Case Is = 151 '(0097H)
        GetFormatTag = "ZyXEL ADPCM"
    Case Is = 152 '(0098H)
        GetFormatTag = "Philips LPCBB"
    Case Is = 153 '(0099H)
        GetFormatTag = "Packed"
    Case Is = 255 '(00FFH)
        GetFormatTag = "AAC"
    Case Is = 256 '(0100H)
        GetFormatTag = "Rhetorex ADPCM"
    Case Is = 257 '(0101H)
        GetFormatTag = "IBM mu-law"
    Case Is = 258 '(0102H)
        GetFormatTag = "IBM A-law"
    Case Is = 259 '(0103H)
        GetFormatTag = "IBM AVC Adaptive Differential Pulse Code Modulation (ADPCM)"
    Case Is = 273 '(0111H)
        GetFormatTag = "Vivo G.723"
    Case Is = 274 '(0112H)
        GetFormatTag = "Vivo Siren"
    Case Is = 291 '(0123H)
        GetFormatTag = "Digital G.723"
    Case Is = 293 '(0125H)
        GetFormatTag = "Sanyo LD ADPCM"
    Case Is = 304 '(0130H)
        GetFormatTag = "Sipro Lab Telecom ACELP NET"
    Case Is = 305 '(0131H)
        GetFormatTag = "Sipro Lab Telecom ACELP 4800"
    Case Is = 306 '(0132H)
        GetFormatTag = "Sipro Lab Telecom ACELP 8V3"
    Case Is = 307 '(0133H)
        GetFormatTag = "Sipro Lab Telecom G.729"
    Case Is = 308 '(0134H)
        GetFormatTag = "Sipro Lab Telecom G.729A"
    Case Is = 309 '(0135H)
        GetFormatTag = "Sipro Lab Telecom Kelvin"
    Case Is = 320 '(0140H)
        GetFormatTag = "Windows Media Video V8"
    Case Is = 336 '(0150H)
        GetFormatTag = "Qualcomm PureVoice"
    Case Is = 337 '(0151H)
        GetFormatTag = "Qualcomm HalfRate"
    Case Is = 341 '(0155H)
        GetFormatTag = "Ring Zero Systems TUB GSM"
    Case Is = 352 '(0160H)
        GetFormatTag = "Microsoft Audio 1"
    Case Is = 353 '(0161H)
        GetFormatTag = "Windows Media Audio V7 / V8 / V9"
    Case Is = 354 '(0162H)
        GetFormatTag = "Windows Media Audio Professional V9"
    Case Is = 355 '(0163H)
        GetFormatTag = "Windows Media Audio Lossless V9"
    Case Is = 512 '(0200H)
        GetFormatTag = "Creative Labs ADPCM"
    Case Is = 514 '(0202H)
        GetFormatTag = "Creative Labs Fastspeech8"
    Case Is = 515 '(0203H)
        GetFormatTag = "Creative Labs Fastspeech10"
    Case Is = 528 '(0210H)
        GetFormatTag = "UHER Informatic GmbH ADPCM"
    Case Is = 544 '(0220H)
        GetFormatTag = "Quarterdeck"
    Case Is = 560 '(0230H)
        GetFormatTag = "I-link Worldwide VC"
    Case Is = 576 '(0240H)
        GetFormatTag = "Aureal RAW Sport"
    Case Is = 592 '(0250H)
        GetFormatTag = "Interactive Products HSX"
    Case Is = 593 '(0251H)
        GetFormatTag = "Interactive Products RPELP"
    Case Is = 608 '(0260H)
        GetFormatTag = "Consistent Software CS2"
    Case Is = 624 '(0270H)
        GetFormatTag = "Sony SCX"
    Case Is = 768 '(0300H)
        GetFormatTag = "Fujitsu FM Towns Snd"
    Case Is = 1024 '(0400H)
        GetFormatTag = "BTV Digital"
    Case Is = 1025 '(0401H)
        GetFormatTag = "Intel Music Coder"
    Case Is = 1104 '(0450H)
        GetFormatTag = "QDesign Music"
    Case Is = 1664 '(0680H)
        GetFormatTag = "VME VMPCM"
    Case Is = 1665 '(0681H)
        GetFormatTag = "AT&T Labs TPC"
    Case Is = 2222 '(08AEH)
        GetFormatTag = "ClearJump LiteWave"
    Case Is = 4096 '(1000H)
        GetFormatTag = "Olivetti GSM"
    Case Is = 4097 '(1001H)
        GetFormatTag = "Olivetti ADPCM"
    Case Is = 4098 '(1002H)
        GetFormatTag = "Olivetti CELP"
    Case Is = 4099 '(1003H)
        GetFormatTag = "Olivetti SBC"
    Case Is = 4100 '(1004H)
        GetFormatTag = "Olivetti OPR"
    Case Is = 4352 '(1100H)
        GetFormatTag = "Lernout & Hauspie Codec (0x1100)"
    Case Is = 4353 '(1101H)
        GetFormatTag = "Lernout & Hauspie CELP Codec (0x1101)"
    Case Is = 4354 '(1102H)
        GetFormatTag = "Lernout & Hauspie SBC Codec (0x1102)"
    Case Is = 4355 '(1103H)
        GetFormatTag = "Lernout & Hauspie SBC Codec (0x1103)"
    Case Is = 4356 '(1104H)
        GetFormatTag = "Lernout & Hauspie SBC Codec (0x1104)"
    Case Is = 5120 '(1400H)
        GetFormatTag = "Norris"
    Case Is = 5121 '(1401H)
        GetFormatTag = "AT&T ISIAudio"
    Case Is = 5376 '(1500H)
        GetFormatTag = "Soundspace Music Compression"
    Case Is = 6172 '(181CH)
        GetFormatTag = "VoxWare RT24 Speech"
    Case Is = 8132 '(1FC4H)
        GetFormatTag = "NCT Soft ALF2CD (www.nctsoft.com)"
    Case Is = 8192 '(2000H)
        GetFormatTag = "Dolby AC3"
    Case Is = 8193 '(2001H)
        GetFormatTag = "Dolby DTS"
    Case Is = 8194 '(2002H)
        GetFormatTag = "WAVE_FORMAT_14_4"
    Case Is = 8195 '(2003H)
        GetFormatTag = "WAVE_FORMAT_28_8"
    Case Is = 8196 '(2004H)
        GetFormatTag = "WAVE_FORMAT_COOK"
    Case Is = 8197 '(2005H)
        GetFormatTag = "WAVE_FORMAT_DNET"
    Case Is = 26447 '(674FH)
        GetFormatTag = "Ogg Vorbis 1"
    Case Is = 26448 '(6750H)
        GetFormatTag = "Ogg Vorbis 2"
    Case Is = 26449 '(6751H)
        GetFormatTag = "Ogg Vorbis 3"
    Case Is = 26479 '(676FH)
        GetFormatTag = "Ogg Vorbis 1+"
    Case Is = 26480 '(6770H)
        GetFormatTag = "Ogg Vorbis 2+"
    Case Is = 26481 '(6771H)
        GetFormatTag = "Ogg Vorbis 3+"
    Case Is = 31265 '(7A21H)
        GetFormatTag = "GSM-AMR (CBR no SID)"
    Case Is = 31266 '(7A22H)
        GetFormatTag = "GSM-AMR (VBR including SID)"
    Case Is = 65534 '(FFFEH)
        GetFormatTag = "WAVE_FORMAT_EXTENSIBLE"
    Case Is = 65535 '(FFFFH)
        GetFormatTag = "WAVE_FORMAT_DEVELOPMENT"
End Select
End Function

Function GUIDStrToHexByteString(GUIDString As String) As String
'Microsoft defines these 16-byte (128-bit) GUIDs as:
'first 4 bytes are in little-endian order
'next 2 bytes are appended in little-endian order
'next 2 bytes are appended in little-endian order
'next 2 bytes are appended in big-endian order
'next 6 bytes are appended in big-endian order
'AaBbCcDd-EeFf-GgHh-IiJj-KkLlMmNnOoPp is stored as this 16-byte string:
'$Dd $Cc $Bb $Aa $Ff $Ee $Hh $Gg $Ii $Jj $Kk $Ll $Mm $Nn $Oo $Pp
Dim hexByteCharString$(16)
hexByteCharString$(0) = (Mid(GUIDString, 7, 2))
hexByteCharString$(1) = (Mid(GUIDString, 5, 2))
hexByteCharString$(2) = (Mid(GUIDString, 3, 2))
hexByteCharString$(3) = (Mid(GUIDString, 1, 2))
hexByteCharString$(4) = (Mid(GUIDString, 12, 2))
hexByteCharString$(5) = (Mid(GUIDString, 10, 2))
hexByteCharString$(6) = (Mid(GUIDString, 17, 2))
hexByteCharString$(7) = (Mid(GUIDString, 15, 2))
hexByteCharString$(8) = (Mid(GUIDString, 20, 2))
hexByteCharString$(9) = (Mid(GUIDString, 22, 2))
hexByteCharString$(10) = (Mid(GUIDString, 25, 2))
hexByteCharString$(11) = (Mid(GUIDString, 27, 2))
hexByteCharString$(12) = (Mid(GUIDString, 29, 2))
hexByteCharString$(13) = (Mid(GUIDString, 31, 2))
hexByteCharString$(14) = (Mid(GUIDString, 33, 2))
hexByteCharString$(15) = (Mid(GUIDString, 35, 2))
GUIDStrToHexByteString = Join(hexByteCharString$)
End Function

Function GUIDStrToCharString(GUIDString As String) As String
'Microsoft defines these 16-byte (128-bit) GUIDs as:
'first 4 bytes are in little-endian order
'next 2 bytes are appended in little-endian order
'next 2 bytes are appended in little-endian order
'next 2 bytes are appended in big-endian order
'next 6 bytes are appended in big-endian order
'AaBbCcDd-EeFf-GgHh-IiJj-KkLlMmNnOoPp is stored as this 16-byte string:
'$Dd $Cc $Bb $Aa $Ff $Ee $Hh $Gg $Ii $Jj $Kk $Ll $Mm $Nn $Oo $Pp
Dim hexByteCharString$(16)
hexByteCharString$(0) = Chr(Val("&H" & (Mid(GUIDString, 7, 2))))
hexByteCharString$(1) = Chr(Val("&H" & (Mid(GUIDString, 5, 2))))
hexByteCharString$(2) = Chr(Val("&H" & (Mid(GUIDString, 3, 2))))
hexByteCharString$(3) = Chr(Val("&H" & (Mid(GUIDString, 1, 2))))
hexByteCharString$(4) = Chr(Val("&H" & (Mid(GUIDString, 12, 2))))
hexByteCharString$(5) = Chr(Val("&H" & (Mid(GUIDString, 10, 2))))
hexByteCharString$(6) = Chr(Val("&H" & (Mid(GUIDString, 17, 2))))
hexByteCharString$(7) = Chr(Val("&H" & (Mid(GUIDString, 15, 2))))
hexByteCharString$(8) = Chr(Val("&H" & (Mid(GUIDString, 20, 2))))
hexByteCharString$(9) = Chr(Val("&H" & (Mid(GUIDString, 22, 2))))
hexByteCharString$(10) = Chr(Val("&H" & (Mid(GUIDString, 25, 2))))
hexByteCharString$(11) = Chr(Val("&H" & (Mid(GUIDString, 27, 2))))
hexByteCharString$(12) = Chr(Val("&H" & (Mid(GUIDString, 29, 2))))
hexByteCharString$(13) = Chr(Val("&H" & (Mid(GUIDString, 31, 2))))
hexByteCharString$(14) = Chr(Val("&H" & (Mid(GUIDString, 33, 2))))
hexByteCharString$(15) = Chr(Val("&H" & (Mid(GUIDString, 35, 2))))
GUIDStrToCharString = Join(hexByteCharString$, "")
End Function

Function GUIDStrToGUIDStruct(GUIDString As String) As GuidStruct
'Microsoft defines these 16-byte (128-bit) GUIDs as:
'first 4 bytes are in little-endian order
'next 2 bytes are appended in little-endian order
'next 2 bytes are appended in little-endian order
'next 2 bytes are appended in big-endian order
'next 6 bytes are appended in big-endian order
'AaBbCcDd-EeFf-GgHh-IiJj-KkLlMmNnOoPp is stored as this 16-byte string:
'$Dd $Cc $Bb $Aa $Ff $Ee $Hh $Gg $Ii $Jj $Kk $Ll $Mm $Nn $Oo $Pp
Debug.Print "Converting "; GUIDString
With GUIDStrToGUIDStruct
    .Data1 = Val("&H" & (Mid(GUIDString, 1, 8)))
    .Data2 = Val("&H" & (Mid(GUIDString, 10, 4)))
    .Data3 = Val("&H" & (Mid(GUIDString, 15, 4)))
    .Data4(0) = Val("&H" & (Mid(GUIDString, 20, 2)))
    .Data4(1) = Val("&H" & (Mid(GUIDString, 22, 2)))
    .Data4(2) = Val("&H" & (Mid(GUIDString, 25, 2)))
    .Data4(3) = Val("&H" & (Mid(GUIDString, 27, 2)))
    .Data4(4) = Val("&H" & (Mid(GUIDString, 29, 2)))
    .Data4(5) = Val("&H" & (Mid(GUIDString, 31, 2)))
    .Data4(6) = Val("&H" & (Mid(GUIDString, 33, 2)))
    .Data4(7) = Val("&H" & (Mid(GUIDString, 35, 2)))
End With
End Function

Function BuildGUID(GUIDIn As GuidStruct) As String
'make it look like 75B22630-668E-11CF-A6D9-00AA0062CE6C
Dim byt As Integer
With GUIDIn
    BuildGUID = AddZero(Hex$(.Data1), 8) & "-"
    BuildGUID = BuildGUID & AddZero(Hex$(.Data2), 4) & "-"
    BuildGUID = BuildGUID & AddZero(Hex$(.Data3), 4) & "-"
    'convert each byte to hex and assemble the string
    For byt = 0 To 1
        BuildGUID = BuildGUID & AddZero(Hex$(.Data4(byt)), 2)
    Next byt
    BuildGUID = BuildGUID & "-"
    For byt = 2 To 7
        BuildGUID = BuildGUID & AddZero(Hex$(.Data4(byt)), 2)
    Next byt
End With
strGUID = BuildGUID
'Debug.Print Len(BuildGUID)
'Debug.Print strGUID
'Debug.Print "75B22630-668E-11CF-A6D9-00AA0062CE6C"
End Function

Function GUIDToString(ByRef GUIDIn As GuidStruct) As String
Dim RetBuf As String
Dim GUILen As Long
Const BufLen As Long = 80
'This function puts Brackets around the guid
RetBuf = Space$(BufLen)
GUILen = StringFromGUID2(GUIDIn, RetBuf, BufLen)
If (GUILen) Then
    GUIDToString = StrConv(Left$(RetBuf, (GUILen - 1) * 2), vbFromUnicode)
    GUIDToString = Mid$(GUIDToString, 2, Len(GUIDToString) - 2)
End If
'Debug.Print GUIDToString
End Function

