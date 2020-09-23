VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWMAHeadReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'submitted by Tom Pydeski
'Structure for reading the Windows media audio files in order to edit tags.
'All in VB from scratch, No DLL's or references.
'
'I saw a lot of examples for editing mp3 tags, but not much to read wma files.
'there was one by Somenon at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61254&lngWId=1
'Although the original author's code worked, I wanted a way to read the file
'without parsing through each character looking for a certain string.
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
'--------------------------------------------------------------------------------------
'below from the attached Micro$oft documentation for streaming media
'ASF files are logically composed of three types of top-level objects:
'the Header Object, the Data Object, and the Index Object(s).
'The Header Object is mandatory and must be placed at the beginning of every ASF file.
'The Data Object is also mandatory and must follow the Header Object.
'The Index Object(s) are optional, but they are useful in providing time-based
'random access into ASF files. When present, the Index Object(s) must be the
'last object in the ASF file.
'
'The Header Object may include a number of standard objects including, but not limited to:
'
'   File Properties Object. Contains global file attributes.
'   Stream Properties Object. Defines a digital media stream and its characteristics.
'   Header Extension Object. Allows additional functionality to be added to an ASF file
'    while maintaining backward compatibility.
'   Content Description Object. Contains bibliographic information.
'   Script Command Object. Contains commands that can be executed on the playback timeline.
'   Marker Object. Provides named jump points within a file.
'Note that objects in the Header Object may appear in any order.
'
'To be valid, the Header Object must contain:
'   a File Properties Object
'   a Header Extension Object
'   and at least one Stream Properties Object.
'
'=====================================================================================
'The Header Object is represented using the following structure.
'------------------------------------------------------
'Field name              Field type   Size (bytes)
'------------------------------------------------------
'Object ID                  GUID        16
'Object Size                QWORD        8
'Number of Header Objects   DWORD        4
'Reserved1                  BYTE         1
'Reserved2                  BYTE         1
'
'The fields are defined as follows:
'object ID
'    Specifies the GUID for the Header Object.
'    This field shall be set to ASF_Header_Object.
'object Size
'    Specifies the size of the Header Object.
'    This includes the sum of 24 bytes plus the size of the objects
'    contained in the Header Object. Valid values are at least 30 bytes.
'Number of Header Objects
'    Specifies the number of objects contained within the Header Object,
'    not including this one. In addition, the Header Extension Object
'    is counted as exactly one object regardless of the number of subobjects
'    contained within it.
'Reserved1
'    This field must be set to the value 0x01.
'    ASF parsers may safely ignore this value.
'Reserved2
'    This field must be set to the value 0x02.
'    If the this value is different when read, the application should fail
'    to source the content.
'
'===============================================================================
'1.1 Content Description Object (optional, one only)
'The Content Description Object is represented using the following structure.
'
'------------------------------------------------------
'Field name             Field type  Size (bytes)
'------------------------------------------------------
'Object ID              GUID        16  'ASF_Content_Description_Object
'Object Size            QWORD       8   'Valid values are at least 34 bytes.
'Title Length           WORD        2
'Author Length          WORD        2
'Copyright Length       WORD        2
'Description Length     WORD        2
'Rating Length          WORD        2
'Title                  WCHAR       Varies
'Author                 WCHAR       Varies
'Copyright              WCHAR       Varies
'Description            WCHAR       Varies
'Rating                 WCHAR       Varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Content Description Object. The value
' of this field shall be set to ASF_Content_Description_Object.
'
'Object Size
'Specifies the size, in bytes, of the Content Description Object.
' Valid values are at least 34 bytes.
'
'Title Length
'Specifies the number of bytes that comprise the title information
' stored in the Title field.
'
'Author Length
'Specifies the number of bytes that comprise the author
' information stored in the Author field.
'
'Copyright Length
'Specifies the number of bytes that comprise the copyright
' information stored in the Copyright field.
'
'Description Length
'Specifies the number of bytes that comprise the description
' information stored in the Description field.
'
'Rating Length
'Specifies the number of bytes that comprise the rating information
' stored in the Rating field.
'
'Title
'Specifies an array of WCHARs that contains the title information.
' Whereas it is highly recommended that this string include a nul
'-terminator, the nul-terminator may not be present.
'
'Author
'Specifies an array of WCHARs that contains the author information.
' Whereas it is highly recommended that this string
' include a nul-terminator, the nul-terminator may not be present.
'
'Copyright
'Specifies an array of WCHARs that contains the copyright information.
' Whereas it is highly recommended that this string
' include a nul-terminator, the nul-terminator may not be present.
'
'Description
'Specifies an array of WCHARs that contains the description information.
' Whereas it is highly recommended that this string
' include a nul-terminator, the nul-terminator may not be present.
'
'Rating
'Specifies an array of WCHARs that contains the rating information.
' Whereas it is highly recommended that this string include a nul
'-terminator, the nul-terminator may not be present.
'
'===============================================================================
'3.11    Extended Content Description Object (optional, one only)
'The Extended Content Description Object lets authors record data
'describing the file and its contents that is beyond the standard
'bibliographic information such as title, author, copyright,
'description, or rating information. This information is pertinent
'to the whole file. Each Content Descriptor stored in this object
'uses a name/value pair metaphor.
'
'The Extended Content Description Object is represented using the following structure.
'
'Field name                 Field type  Size (bytes)
'Object ID                      GUID    16
'Object Size                    QWORD   8
'Content Descriptors Count      WORD    2
'Content Descriptors See text   varies
'
'Object ID
'Specifies the GUID for the Extended Content Description Object.
' The value of this field shall be set to ASF_Extended_Content_Description_Object.
'
'Object Size
'Specifies the size, in bytes, of the Extended Content Description Object.
' Valid values are at least 26 bytes.
'
'Content Descriptors Count
'Specifies the number of entries in the Content Descriptors list.
'
'Content Descriptors
'The structure of each Content Descriptor entry is shown in the following table.
'Field Name                 Field Type  Size (bytes)
'Descriptor Name Length         WORD    2
'Descriptor Name                WCHAR   varies
'Descriptor Value Data Type     WORD    2
'Descriptor Value Length        WORD    2
'Descriptor Value             See text  varies
'
'Descriptor Value Data Type
'Specifies the type of data stored in the Descriptor Value field.
'The types are defined in the following table.
'Value   Type               Descriptor value length
'0x0000  Unicode string     varies
'0x0001  BYTE array         varies
'0x0002  BOOL               4
'0x0003  DWORD              4
'0x0004  QWORD              8
'0x0005  WORD               2
'
'===============================================================================
'Description Records
'The structure of each Description Record entry is shown in the following table.
'Field name             Field type  Size (bytes)
'Reserved (Must Be Zero)    WORD    2
'Stream Number              WORD    2
'Name Length                WORD    2
'Data Type                  WORD    2
'Data Length                DWORD   4
'Name                       WCHAR   varies
'Data                       See below
'
'Data Type
'Specifies the type of information being stored. The following values are recognized.
'Value type  Description
'0x0000  Unicode string The data consists of a sequence of Unicode characters.
'0x0001  BYTE array.    The type of data is implementation-specific.
'0x0002  BOOL.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'                       Only 0x0000 or 0x0001 are permitted values.
'0x0003  DWORD.         The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'0x0004  QWORD.         The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'0x0005  WORD.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'
'=====================================================================================
'File Properties Object (mandatory, one only)
'   The File Properties Object defines the global characteristics of
'   the combined digital media streams found within the Data Object.
'
'The File Properties Object is represented using the following structure.
'
'Field name               Field type  Size (bytes)
'Object ID                  GUID        16
'Object Size                QWORD        8
'File ID                    GUID        16
'File Size                  QWORD        8
'Creation Date              QWORD        8
'Data Packets Count         QWORD        8
'Play Duration              QWORD        8
'Send Duration              QWORD        8
'Preroll                    QWORD        8
'Flags                      DWORD        4
'Broadcast                  Flag         1 bit (LSB)
'Seekable                   Flag         1 bit
'Reserved                               30 bits
'Minimum Data Packet Size   DWORD        4
'Maximum Data Packet Size   DWORD        4
'Maximum Bitrate            DWORD        4
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the File Properties Object.
'This field shall be set to ASF_File_Properties_Object.
'
'Object Size
'Specifies the size, in bytes, of the File Properties Object.
'Valid values are at least 104 bytes.
'
'File ID
'Specifies the unique identifier for this file.
'The value shall be regenerated every time the file is modified in any way.
'The value shall be identical to the value of the File ID field of the Data Object.
'
'File Size
'Specifies the size, in bytes, of the entire file.
'The value is invalid if the Broadcast Flag bit in the Flags field is set to 1.
'
'Creation Date
'Specifies the date and time of the initial creation of the file.
'The value is given as the number of 100-nanosecond intervals since January 1, 1601,
' according to Coordinated Universal Time (Greenwich Mean Time).
'The value may be invalid if the Broadcast Flag bit in the Flags field is set to 1.
'
'Data Packets Count
'Specifies the number of Data Packet entries that exist within the Data Object.
'The value is invalid if the Broadcast Flag bit in the Flags field is set to 1.
'
'Play Duration
'Specifies the time needed to play the file in 100-nanosecond units.
' This value should include the duration (estimated, if an exact value is unavailable)
'of the the last media object in the presentation.
'The value is invalid if the Broadcast Flag bit in the Flags field is set to 1.
'
'Send Duration
'Specifies the time needed to send the file in 100-nanosecond units.
'This value should include the duration of the last packet in the content.
'The value is invalid if the Broadcast Flag bit in the Flags field is set to 1.
'Players can ignore this value.
'
'Preroll
'Specifies the amount of time to buffer data before starting to play the file, in millisecond units.
'If this value is nonzero, the Play Duration field and all of the payload Presentation Time fields have
'been offset by this amount. Therefore, player software must subtract the value in the preroll field
'from the play duration and presentation times to calculate their actual values.
'It follows that all payload Presentation Time fields need to be at least this value.
'
'Flags
'The flags are stored in Least Significant Byte (LSB) order.
'
'Broadcast Flag (bit 0)
'Specifies, if set, that a file is in the process of being created
'(for example, for recording applications),
'and thus that various values stored in the header objects are invalid.
'It is highly recommended that post-processing be performed
'to remove this condition at the earliest opportunity.
'
'Seekable Flag (bit 1)
'Specifies, if set, that a file is seekable.
'Note that for files containing a single audio stream and a Minimum Data Packet Size field
'equal to the Maximum Data Packet Size field, this flag shall always be set to 1.
'For files containing a single audio stream and a video stream or mutually exclusive
'video streams, this flag is only set to 1 if the file contains a matching
'Simple Index Object for each regular video stream
' (that is, video streams that are not hidden according to the method described in section 8.2.2).
'
'Reserved (bits 2 - 31)
'Remaining 30 reserved flags shall be set to 0.
'
'Minimum Data Packet Size
'Specifies the minimum Data Packet size in bytes.
' In general, the value is invalid if the Broadcast Flag bit in the Flags field is set to 1.
' However, for the purposes of this specification, the values for the Minimum Data Packet Size
' and Maximum Data Packet Size fields shall be set to the same value, and this value
'should be set to the packet size, even when the Broadcast Flag in the Flags field is set to 1.
'
'Maximum Data Packet Size
'Specifies the maximum Data Packet size in bytes.
'In general, the value is invalid if the Broadcast Flag bit in the Flags field is set to 1.
'However, for the purposes of this specification, the values of the Minimum Data Packet Size
'and Maximum Data Packet Size fields shall be set to the same value, and this value
'should be set to the packet size, even when the Broadcast Flag field is set to 1.
'
'Maximum Bitrate
'Specifies the maximum instantaneous bit rate in bits per second for the entire file.
' This shall equal the sum of the bit rates of the individual digital media streams.
'It shall be noted that the digital media stream includes ASF data packetization overhead
'as well as digital media data in payloads. Only those streams that have a free-standing
'Stream Properties Object in the header shall have their bit rates included in the sum;
'streams whose Stream Properties Object exists as part of an
'Extended Stream Properties Object in the Header Extension Object shall not have
'their bit rates included in this sum, except when this value would otherwise be 0.
'=====================================================================================
'3.12  Stream Bitrate Properties Object (optional but recommended, one only)
'The Stream Bitrate Properties Object defines the average bit rate
' of each digital media stream.
' It is represented using the following structure.
'
'Field name             Field type  Size (bytes)
'---------------------------------------------------------
'Object ID              GUID        16
'Object Size            QWORD       8
'Bitrate Records Count  WORD        2
'Bitrate Records        See below   varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Stream Bitrate Properties Object.
' The value of this field shall be set to ASF_Stream_Bitrate_Properties_Object.
'
'Object Size
'Specifies the size, in bytes, of the Stream Bitrate Properties Object.
' Valid values are at least 26 bytes.
'
'Bitrate Records Count
'Specifies the number of records listed in the Bitrate Records field.
'
'Bitrate Records
'The structure of each Bitrate Record entry is shown in the following table.
'Field name         Field type  Size (bits)
'------------------------------------------------------
'Flags              WORD    2
'  -Stream Number           -7 bits (LSB)
'  -Reserved                -9 bits
'Average Bitrate    DWORD   4
'
'The fields are defined as follows:
'Flags
'The flags are stored in LSB order.
'Stream Number (bits 0-6)
'Specifies the number of this stream described by this record.
' 0 is an invalid stream. Valid values are between 1 and 127.
'
'Reserved (bits 7-15)
'These bits are reserved and should be set to 0.
'
'Average Bitrate
'Specifies the average bit rate of the stream in bits per second.
' This value should include an estimate of ASF packet and payload
' overhead associated with this stream.    Stream Bitrate Properties Object (optional but recommended, one only)
'=====================================================================================
'1.1    Header Extension Object (mandatory, one only)
'The Header Extension Object allows additional functionality to be added to an ASF file while
'maintaining backward compatibility. The Header Extension Object is a container containing 0 or
'more additional extended header objects. Extended header objects should conform to the ASF
'Object Structure (per section 2.1). For more information, see section 8.2.1.
'
'The Header Extension Object is represented using the following structure.
'
'Field Name                 Field Type  Size (bits)
'Object ID                      GUID    16
'Object Size                    QWORD   8
'Reserved Field 1               GUID    16
'Reserved Field 2               WORD    2
'Header Extension Data Size     DWORD   4
'Header Extension Data          BYTE    varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Header Extension Object.
' The value of this field shall be set to ASF_Header_Extension_Object.
'
'Object Size
'Specifies the size, in bytes, of the Header Extension Object.
' The value of this field shall be set to 46 bytes.
'
'Reserved Field 1
'This value shall be set to ASF_Reserved_1 as defined in section 10.6.
'
'Reserved Field 2
'This value should be set to 6.
'
'Header Extension Data Size
'Specifies the number of bytes stored in the Header Extension Data field.
' This value may be 0 bytes or 24 bytes and larger. It should also be equal
' to the Object Size field minus 46 bytes.

'Header Extension Data
'Specifies an array of bytes containing additional extended header objects.
' This data should be interpreted as 0 or more extended header objects stored
' consecutively within the array of bytes. No empty space, padding, leading,
' or trailing bytes are allowed.
'
'=====================================================================================
'3.5 Codec List Object (optional, one only)
'The Codec List Object provides user-friendly information about the codecs and
'formats used to encode the content found in the ASF file.
'The Codec List Object is represented using the following structure.
'
'Field name              Field type  Size (bytes)
'Object ID               GUID        16
'Object Size             QWORD       8
'Reserved                GUID        16
'Codec Entries Count     DWORD       4
'Codec Entries           See below   varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Codec List Object.
' The value of this field shall be set to ASF_Codec_List_Object.
'
'Object Size
'Specifies the size, in bytes, of the Codec List Object.
'Valid values are at least 44 bytes.
'
'Reserved
'This field shall be set to ASF_Reserved_2 as defined in section 10.8.
'
'Codec Entries Count
'Specifies the number of entries listed in the Codec Entries field.
'
'Codec Entries
'Codec Entries are described in the following table.
'Field Name              Field Type     Size (bytes)
'------------------------------------------------------
'Type                       WORD        2
'Codec Name Length          WORD        2
'Codec Name                 WCHAR       varies
'Codec Description Length   WORD        2
'Codec Description          WCHAR       varies
'Codec Information Length   WORD        2
'Codec Information          BYTE        varies
'
'The fields are defined as follows:
'Type
'Specifies the type of the codec used.
'Use one of the values in the following table.
'   Values  Meaning
'   0x0001  Video codec
'   0x0002  Audio codec
'   0xFFFF  Unknown codec
'
'Codec Name Length
'Specifies the number of Unicode characters stored in the Codec Name field.
'
'Codec Name
'Specifies an array of Unicode characters that contains the name of the codec
'used to create the content.
'
'Codec Description Length
'Specifies the number of Unicode characters stored in the Codec Description field.
'
'Codec Description
'Specifies an array of Unicode characters that contains the description of
'the format used to create the content.
'
'Codec Information Length
'Specifies the number of bytes stored in the Codec Information field.
'
'Codec Information
'Specifies an opaque array of information bytes about the codec used to
'create the content. The meaning of these bytes is determined by the codec.
'
'======================================================================================
'The Stream Properties Object is represented using the following structure.
'
'Field Name                 Field Type  Size (bits)
'Object ID                      GUID    128
'Object Size                    QWORD   64
'Stream Type                    GUID    128
'Error Correction Type          GUID    128
'Time Offset                    QWORD   64
'Type-Specific Data Length      DWORD   32
'Error Correction Data Length   DWORD   32
'Flags                          WORD    16
'   Stream Number                  7 (LSB)
'   Reserved                       8
'   Encrypted                      Content Flag     1
'Reserved                       DWORD   32
'Type-Specific                  Data BYTE    varies
'Error Correction               Data  BYTE    varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Stream Properties Object.
' The value of this field shall be set to ASF_Stream_Properties_Object.
'
'Object Size
'Specifies the size, in bytes, of the Stream Properties Object.
' Valid values are at least 78 bytes.
'
'Stream Type
'Specifies the type of the stream (for example, audio, video, and so on).
' Use one of the GUIDs defined in section 10.4. Any streams with unrecognized Stream
' Type values should be ignored.
'
'Error Correction Type
'Specifies the error correction type used by this digital media stream.
' For streams other than audio, this value should be set to ASF_No_Error_Correction.
' For audio streams, this value should be set to ASF_Audio_Spread.
' The possible GUIDs are defined in section 10.5, and the standard error
' correction schemes are described in section 9.1.
'
'Time Offset
'Specifies the presentation time offset of the stream in 100-nanosecond units.
' The value of this field is added to all of the timestamps of the samples in the stream.
' This value shall be equal to the send time of the first interleaved packet in the data section.
' The value of this field is typically 0. It is non-zero in the case when an ASF file
' is edited and it is not possible for the editor to change the presentation times
' and send times of ASF packets. Note that if more than one stream is present in
' an ASF file the offset values of all stream properties objects must be equal.
'
'Type-Specific Data Length
'Specifies the number of bytes in the Type-Specific Data field.
'
'Error Correction Data Length
'Specifies the number of bytes in the Error Correction Data field.
'
'Flags
'The flags are stored in LSB order.
'
'Stream Number (bits 0-6)
'Specifies the number of this stream. 0 is an invalid stream.
' Valid values are between 1 and 127. The numbers assigned to streams
' in an ASF presentation may be any combination of unique values;
' parsing logic must not assume that streams are numbered sequentially.
'Reserved (bits 7-14) These bits are reserved and should be set to 0.
'Encrypted Content Flag (bit 15)
'Specifies, if set, that the data contained in this stream is encrypted
'and will be unreadable unless there is a way to decrypt the stream.
'
'Reserved
'This field is reserved and should be set to 0.
'
'Type-Specific Data
'Specifies type-specific format data.
' The structure for the Type-Specific Data field is determined by the
' value stored in the Stream Type field. The structure for the
' Type-Specific Data field for standard ASF media types is detailed in section 9.
'
'Error Correction Data
'Specifies data specific to the error correction type.
' The structure for the Error Correction Data field is determined by
' the value stored in the Error Correction Type field.
' For example, an audio data stream might need to know how codec chunks
' were redistributed, or it might need a sample of encoded silence. For
' detailed information, see section 9.1.
'
'======================================================================================
' Audio media type - ASF_Audio_Media
'When the Stream Type of the Stream Properties Object has the value ASF_Audio_Media,
'the ASF audio media type that populates the Type-Specific Data field of the
'Stream Properties Object is represented using the following structure
'(the WAVEFORMATEX structure).
'
'Field name              Field type  Size (bytes)
'-----------------------------------------------------------------------------------
'Codec ID / Format Tag      WORD        2
'Number of Channels         WORD        2
'Samples Per Second         DWORD       4
'Avg # of Bytes Per Sec     DWORD       4
'Block Alignment            WORD        2
'Bits Per Sample            WORD        2
'Codec Specific Data Size   WORD        2
'Codec Specific Data        BYTE     varies
'

'below another sample
'is it in 1600h range?
'61 01 02 00 22 56 00 00 A6 0F 00 00 E8 02 10 00
'0A 00 00 44 00 00 17 00 41 17 00 00 01 E8 02 E8
'02 01 00 00
'
'my file is this:
'61 01 02 00 44 AC 00 00 22 4E 00 00 07 1D 10 00
'0A 00 00 88 00 00 0F 00 1D 74 00 00 01 07 1D 07
'1D 01
'
'Codec ID : 01 61
'Number of Channels : 00 02
'Samples Per Second : 00 00 56 22 = 22050
'Average Number of Bytes Per Second: 00 00 of a6 = 4006
'Block Alignment: 02 e8
'Bits Per Sample: 00 10
'Codec Specific Data Size: 00 0A
'Codec Specific Data:  00 44 00 00 17 00 41 17 00 00
'
'The fields are defined as follows:
'
'Codec ID / Format Tag
'Specifies the unique ID of the codec used to encode the audio data.
' There is a registration procedure for new codecs.
' Defined as the wFormatTag field of a WAVEFORMATEX structure.
'
'Number of Channels
'Specifies the number of audio channels. Monaural data uses one channel
' and stereo data uses two channels. 5.1 audio uses six channels.
' Defined as the nChannels field of a WAVEFORMATEX structure.
'
'Samples Per Second
'Specifies a value in Hertz (cycles per second) that represents the
' sampling rate of the audio stream.
' Defined as the nSamplesPerSec field of a WAVEFORMATEX structure.
'
'Average Number of Bytes Per Second
'Specifies the average number of bytes per second of the audio stream.
' Defined as the nAvgBytesPerSec field of a WAVEFORMATEX structure.
'
'Block Alignment
'Specifies the block alignment, or block size, in bytes of the audio codec.
' Defined as the nBlockAlign field of a WAVEFORMATEX structure.
'
'Bits per Sample
'Specifies the number of bits per sample of monaural data.
' Defined as the wBitsPerSample field of a WAVEFORMATEX structure.
'
'Codec Specific Data Size
'Specifies the size, in bytes, of the Codec Specific Data buffer.
' Defined as the cbSize field of a WAVEFORMATEX structure.
' This value should be 0 when Codec ID is 1 (WAVE_FORMAT_PCM).
'
'Codec Specific Data
'Specifies an array of codec-specific data bytes.
'
'======================================================================================
'Compatibility Object  (26 bytes)
'Object ID,26F18B5D-4584-47EC-9F5F-0E651F0452C9,
'Object Size,26  ( 0x1A ),
'Profile,2,
'Mode,1,
'
'4.12   Object (optional, only 1)
'The Compatibility Object is reserved for future use.
'It is represented using the following structure.
'Compatibility Field Name   Field Type  Size (bytes)
'Object ID                  GUID        16
'Object Size                QWORD       8
'Profile                    BYTE        1
'Mode                       BYTE        1
'
'Object ID
'Specifies the GUID for the Compatibility Object.
'The value of this field shall be set to ASF_Compatibility_Object.
'
'Object Size
'Specifies the size, in bytes, of the Compatibility Object.
'Valid values are at least 28 bytes.
'
'Profile
'This field is reserved and should be set to 2.
'
'Mode
'This field is reserved and should be set to 1.
'
'======================================================================================
'4.7    Metadata Object (optional, 0 or 1)
'The Metadata Object permits authors to store stream-based metadata in a file.
'
'This object supports the same types of metadata information as the
'Extended Content Description Object except that it also allows a
'stream number to be specified. For more information about what
'types of attributes belong in the Metadata Object, see section 7.4.
'
'The Metadata Object is represented using the following structure.
'
'Field name                 Field type  Size (bytes)
'Object ID                  GUID        16
'Object Size                QWORD       8
'Description Records Count  WORD        2
'Description Records        See below   varies
'
'The fields are defined as follows:
'Object ID
'Specifies the GUID for the Metadata Object.
'The value of this field shall be set to ASF_Metadata_Object.
'
'Object Size
'Specifies the size, in bytes, of the Metadata Object.
'Valid values are larger than 26 bytes.
'
'Description Records Count
'Specifies the number of entries in the Description Records list.
'
'Description Records
'The structure of each Description Record entry is shown in the following table.
'Field name                 Field type  Size (bytes)
'Reserved (Must Be Zero)    WORD        2
'Stream Number              WORD        2
'Name Length                WORD        2
'Data Type                  WORD        2
'Data Length                DWORD       4
'Name                       WCHAR     varies
'Data                       See below
'
'The fields are defined as follows:
'Reserved (Must Be Zero)
' This field must contain the value 0.
'
'Stream Number
'Specifies whether the entry applies to a specific digital media stream
' or whether it applies to the whole file.
' A value of 0 in this field indicates that it applies to the whole file;
' otherwise, the entry applies only to the indicated stream number
' and must be between 1 and 127.
'
'Name Length
'Specifies the number of valid bytes stored in the Name field.
' Valid values are even numbers.
'
'Data Type
'Specifies the type of information being stored. (Standard).
'
'Data Length
'Specifies the length in bytes of the Data field.
' Valid values are less than 65536.
'
'Name
'Specifies the name that uniquely identifies the attribute being described.
' Names are case-sensitive.
'
'Data
'Specifies the actual metadata being stored.
' The Data field should be interpreted according to the value stored in the Data Type field.
'======================================================================================
'
'1.1    Padding Object (optional, 0 to many)
'The Padding Object is a dummy object that is used to pad the size of the Header Object.
'This object enables the size of any object stored in the Header Object to grow or shrink
'without having to rewrite the entire Data Object and Index Object sections of the ASF file.
'For instance, if entries in the Content Description Object or Extended Content Description
'Object need to be removed or shortened, the size of the Padding Object can be increased to
'compensate for the reduction in size of the Content Description Object. The ASF file can then
'be updated by overwriting the previous Header Object with the edited Header Object of identical size,
' without having to move or rewrite the data contained in the Data Object.
'Playback applications shall simply ignore Padding Objects.
'
'The Padding Object is represented using the following structure.
'Field name     Field type  Size (bytes)
'------------------------------------------------
'Object ID      GUID        16
'Object Size    QWORD       8
'Padding Data   BYTE        varies 'Object Size - 24 bytes
'
'The fields are defined as follows:
'Object ID
'Specifies the GUID for the Padding Object.
' The value of this field shall be set to ASF_Padding_Object.
'
'Object Size
'Specifies the size, in bytes, of the Padding Object.
' Valid values are at least 24 bytes.
'
'Padding Data
'Specifies an array of bytes used for padding purposes.
' The length of the padding object data is calculated as Object Size - 24 bytes
' ( sizeof( Object ID ) + sizeof( Object Size ) ). As the name implies, the
' Padding Data field contains padding instead of meaningful data and shall always be ignored.
'======================================================================================
'
'1.1    Extended Stream Properties Object (optional, 1 per media stream)
'The Extended Stream Properties Object defines additional optional properties and characteristics
'of a digital media stream that are not described in the Stream Properties Object.
'
'Typically, the basic Stream Properties Object is present in the Header Object, and the
'Extended Stream Properties Object is present in the Header Extension Object.
'Sometimes, however, the Stream Properties Object for a stream may be embedded inside
'the Extended Stream Properties Object for that stream. This approach facilitates the creation
'of backward-compatible content and is appropriate for some of the scenarios discussed in section 8.
'
'This object has an optional provision to include application-specific or implementation-specific
'data attached to the payloads of each digital media sample stored within a Data Packet.
'This data can be looked at as digital media sample properties and is stored in the Replicated Data
'field of a payload header. The Payload Extension Systems fields of the Extended Stream Properties Object
'describes what this data is and is necessary for that data to be parsed, if present.
'
'The Extended Stream Properties Object is represented using the following structure.
'
'Field name                     Field type  Size (bytes)
'-------------------------------------------------------------
'Object ID                          GUID    16
'Object Size                        QWORD   8
'Start Time                         QWORD   8
'End Time                           QWORD   8
'Data Bitrate                       DWORD   4
'Buffer Size                        DWORD   4
'Initial Buffer Fullness            DWORD   4
'Alternate Data Bitrate             DWORD   4
'Alternate Buffer Size              DWORD   4
'Alternate Initial Buffer Fullness  DWORD   4
'Maximum Object Size                DWORD   4
'Flags                              DWORD   4
'   Reliable Flag (LSB)
'   Seekable Flag
'   No Cleanpoints Flag
'   Resend Live Cleanpoints Flag
'   Reserved Flags                          28 bits
'Stream Number                      WORD        2
'Stream Language ID Index           WORD        2
'Average Time Per Frame             QWORD       8
'Stream Name Count                  WORD        2
'Payload Extension System Count     WORD        2
'Stream Names                   See below   varies
'Payload Extension Systems      See below   varies
'Stream Properties Object       See below   varies
'
'The fields are defined as follows:
'Object ID
'Specifies the GUID for the Extended Stream Properties Object.
' The value of this field shall be set to ASF_Extended_Stream_Properties_Object.
'
'Object Size
'Specifies the size, in bytes, of the Extended Stream Properties Object.
' Valid values are at least 88 bytes. If there is a Stream Properties Object embedded
' in this object, then the size of the Stream Properties Object must be included
' in this value.
'
'Start Time
'Specifies the presentation time of the first object, indicating where this digital
' media stream starts within the context of the timeline of the ASF file as a whole.
' This time value corresponds to presentation times as they appear in the data packets
' (adjusted by the preroll).
' This field is given in units of milliseconds and can optionally be set to 0,
' in which case it will be ignored.
'
'End Time
'Specifies the presentation time of the last object plus the duration of play, indicating
'where this digital media stream ends within the context of the timeline of the ASF file
' as a whole.
' This time value corresponds to presentation times as they appear in the data packets
' (adjusted by the preroll).
' This field is given in units of milliseconds and can optionally be set to 0,
' in which case it will be ignored.
'
'Data Bitrate
'Specifies the leak rate R, in bits per second, of a leaky bucket that contains the
'data portion of the stream without overflowing, excluding all ASF Data Packet overhead.
'The size of the leaky bucket is specified by the value of the Buffer Size field.
'This field must be non-zero. For more information about leaky buckets, see section 7.1.
'
'Buffer Size
'Specifies the size B, in milliseconds, of the leaky bucket used in the
'Data Bitrate definition.
'
'Initial Buffer Fullness
'Specifies the initial fullness, in milliseconds, of the leaky bucket used in the
' Data Bitrate definition. This is the fullness of the buffer at the instant before
' the first bit in the stream is dumped into the bucket. Typically, this value is set to 0.
' This value shall not exceed the value in the Buffer Size field.
'
'Alternate Data Bitrate
'Specifies the leak rate RAlt, in bits per second, of a leaky bucket that contains the
' data portion of the stream without overflowing, excluding all ASF Data Packet overhead.
' The size of the leaky bucket is specified by the value of the Alternate Buffer Size field.
' This value is relevant in most scenarios where the bit rate is not exactly constant,
' but it is especially useful for streams that have highly variable bit rates.
' This field can optionally be set to the same value as the Data Bitrate field.
'
'Alternate Buffer Size
'Specifies the size BAlt, in milliseconds, of the leaky bucket used in the Alternate Data
' Bitrate definition. This value is relevant in most scenarios where the bit rate is not
' constant, but it is especially useful for streams that have highly variable bit rates.
'This field can optionally be set to the same value as the Buffer Size field.
'
'Alternate Initial Buffer Fullness
'Specifies the initial fullness, in milliseconds, of the leaky bucket used in the
' Alternate Data Bitrate definition.  This is the fullness of the buffer at the instant
' before the first bit in the stream is dumped into the bucket.
' Typically, this value is set to 0.
' This value shall not exceed the value of the Alternate Buffer Size field.
'
'Maximum Object Size
'Specifies the maximum size of the largest sample stored in the data packets for a stream.
'A value of 0 means “unknown”.
'
'Flags
'The flags are stored in LSB order.
'   Reliable Flag
'   Specifies, if set, that this digital media stream, if sent over a network, must
'    be carried over a reliable data communications transport mechanism.
'    This should be set for streams that cannot recover after a lost media object.
'
'   Seekable Flag
'   This flag should be set only if the stream is seekable, either by using an index
'    object or by estimating according to bit rate (as can sometimes be done with audio).
'    This flag pertains to this stream only rather than to the entire file.
'
'   No Cleanpoint Flag
'   Specifies, if set, that the stream does not contain any cleanpoints.
'    A cleanpoint is any point at which playback could begin without having seen
'    the previous media objects.  For streams that use key frames, the key frames
'    would be the cleanpoints.
'
'   Resend Live Cleanpoints Flag
'   Specifies, if set, that when a stream is joined in mid-transmission, all information
'    from the most recent cleanpoint up to the current time should be sent before normal
'    streaming begins at the current time.  The default behavior(when this flag is not set)
'    is to send only the data starting at the current time.
'    This flag should only be set for streams that are coming from a “live” source.
'
'   Reserved Flags
'   Specifies reserved flags. Shall be set to 0.
'
'AverageTime Per Frame
'Specifies the average time duration, measured in 100-nanosecond units, of each frame.
' This number should be rounded to the nearest integer.
' This field can optionally be set to 0 if the average time per frame is unknown or unimportant.
' It is recommended that this field be set for video.
'
'Stream Number
'Specifies the number of this stream.
' 0 is an invalid stream number (that is, other Header Objects use stream number 0 to
' refer to the entire file as a whole rather than to a specific media stream within the file).
' Valid values are between 1 and 127.
'
'Stream Language ID Index
'Specifies the language, if any, which the content of the stream uses or assumes.
' Refer to the Language List Object description for the details concerning how the
' Stream Language ID Index and Language ID Index fields should be used.
' Note that this is an index into the languages listed in the
' Language List Object rather than a language identifier.
'
'Stream Name Count
'Specifies how many Stream Names are present. Each stream name instance is potentially
' localized into a specific language. The Language ID Index field indicates the language
' in which the Stream Name has been written in Unicode characters.
' This shall be set to 0 if there are no Stream Names, in which case the
' Stream Names field should be omitted.
'
'Payload Extension System Count
'Specifies how many Payload Extension Systems there are for this stream.
' This shall be set to 0 if there are no Payload Extension Systems, in which
' case the Payload Extension Systems field should be omitted.
'
'Stream Names
'The structure of each Stream Name entry is shown in the following table.
'Field name             Field type  Size (bytes)
'-----------------------------------------------------------------------------
'Language ID Index      WORD        2
'Stream Name Length     WORD        2
'Stream Name            WCHAR   varies
'
'where the Stream Name Length field indicates the number of valid bytes that are found within the Stream Name field.
'
'Payload Extension Systems
'Payload extensions provide a way for content creators to specify kinds of data that will appear in the payload header
' for every payload from this stream. This system is used when stream properties must be conveyed at the media
' object level. The Replicated Data bytes in the payload header will contain these properties in the order in which the
' Payload Extension Systems appear in this object. A Payload Extension System must appear in the Extended Stream
' Properties Object for each type of per-media-object properties that will appear with the payloads for this stream.
' For more information, see section 7.3.
'Payload extension systems are described in the following table.
'Field name                     Field type  Size (bytes)
'--------------------------------------------------------------------------------------
'Extension System ID            GUID        16
'Extension Data Size            WORD        2
'Extension System Info Length   DWORD       4
'Extension System Info          BYTE        varies
'
'The fields are defined as follows:
'Extension System ID
'Specifies a unique identifier for the extension system. Five standard GUIDs are defined in section 10.13,
' although custom extension system types can be defined as well.
'
'Extension Data Size
'Specifies the fixed size of the extension data for this system that will appear in the replicated data alongside
' every payload for this stream. If this extension system uses variable-size data, then this should be set to 0xffff.
' Note, however, that replicated data length is limited to 255 bytes, which limits the total size of all extension
' systems for a particular stream.
'
'Extension System Info Length
'Specifies the length of the Extension System Info field. This field shall be set to 0 if there is no value in the
' Extension System Info field.
'
'Extension System Info
'Specifies additional information to describe this extension system (optional).
'
'Stream Properties Object
'Specifies an optional Stream Properties Object.
'For more information, see the beginning of this section.
'
'--------------------------------------------------------------------------------------
'4.8 Metadata Library Object (optional, 0 or 1)
'The Metadata Library Object lets authors store stream-based, language-attributed,
'multiply defined, and large metadata attributes in a file.
'This object supports the same types of metadata as the Metadata Object,
'as well as attributes with language IDs, attributes that are defined more than once,
'large attributes, and attributes with the GUID data type
'For more information about what types of attributes belong in the Metadata Library Object,
'see section 7.4.
'
'The ASF Metadata Library Object is represented using the following structure.
'
'Field Name               Field Type  Size (bytes)
'--------------------------------------------------
'Object ID                  GUID        16  ASF_Metadata_Library_Object
'Object Size                QWORD       8   at least 26 bytes
'Description Records Count  WORD        2
'Description Records        See below   varies
'
'The fields are defined as follows:
'object ID
'Specifies the GUID for the Metadata Library Object.
' The value of this field shall be set to ASF_Metadata_Library_Object.
'
'object Size
'Specifies the size, in bytes, of the Metadata Library Object.
' Valid values are at least 26 bytes.
'
'Description Records Count
'Specifies the number of entries in the Description Records list.
'
'Description Records
'The structure of each Description Record entry is shown in the following table.
'Field Name          Field Type  Size (bytes)
'--------------------------------------------------
'Language List Index    WORD    2
'Stream Number          WORD    2
'Name Length            WORD    2
'Data Type              WORD    2
'Data Length            DWORD   4
'Name                   WCHAR   varies
'Data               See below   varies
'
'The fields are defined as follows:
'Language List Index
'Specifies the index into the Language List Object that identifies the language
' of this attribute. If there is no Language List Object present, this field must be zero.
'
'Stream Number
'Specifies whether the entry applies to a specific digital media stream or whether it
'applies to the whole file. A value of 0 in this field indicates that it applies to the
'whole file; otherwise, the entry applies only to the indicated stream number.
' Valid values are between 1 and 127.
'
'Name Length
'Specifies the number of valid bytes stored in the Name field.
' Valid values are even numbers.
'
'Data Type
'Specifies the type of information being stored.
'Metadata Library data type
'   0x0000  Unicode string. The data consists of a sequence of Unicode characters.
'   0x0001  BYTE array. The type of the data is implementation-specific.
'   0x0002  BOOL. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'                 Only 0x0000 or 0x0001 are permitted values.
'   0x0003  DWORD. The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'   0x0004  QWORD. The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'   0x0005  WORD. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'   0x0006  GUID. The data is 16 bytes long and should be interpreted as a 128-bit GUID.
'
'
'Data Length
'Specifies the length, in bytes, of the Data field.
'
'Name
'Specifies the name that identifies the attribute being described.
'
'Data
'Specifies the actual metadata being stored.
' The Data field should be interpreted according to the value stored in the Data Type field.
'
'--------------------------------------------------------------------------------------
'5.1 ASF Data Object (mandatory, one only)
'The Data Object is represented using the following structure.
'
'Field name           Field type    Size (bits)
'-----------------------------------------------------------
'Object ID              GUID        16 ' ASF_Data_Object
'Object Size            QWORD       8 'at least 50 bytes
'File ID                GUID        16
'Total Data Packets     QWORD       8
'Reserved               WORD        2 'set to 0x0101. (257D)
'Data Packets    See section 5.2
'varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Data Object.
'The value of this field shall be set to ASF_Data_Object.
'
'Object Size
'Specifies the size of the Data Object.
'Valid values are at least 50 bytes. Note that if the Broadcast Flag bit
'of the Flags field is set on the File Properties Object, then the
'Object Size field may be 0. This is a special case that indicates
'the size of the Data Object is unknown. It is not valid to have a value of 0
'for the Object Size field if the Broadcast Flag is not set.
'
'File ID
'Specifies the unique identifier for this ASF file.
'The value of this field shall be changed every time the file is modified in any way.
'The value of this field shall be identical to the value of the File ID field
'of the Header Object.
'
'Total Data Packets
'Specifies the number of ASF Data Packet entries that exist within the Data Object.
'It must be equal to the Data Packet Count field in the File Properties Object (section 3.2).
'The value of this field is invalid if the Broadcast Flag field of the
'File Properties Object is set to 1.
'
'Reserved
'Specifies a reserved field. The value of this field shall set to 0x0101.
'
'Data Packets
'Specifies a list of Data Packets, as defined in section 5.2.
'
'This structure is immediately followed by one or more Data Packets.

'======================================================================================
'Top-level ASF object GUIDS
'The following table contains the names and values of top-level ASF object GUIDs.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Header_Object                          75B22630-668E-11CF-A6D9-00AA0062CE6C
'ASF_Data_Object                            75B22636-668E-11CF-A6D9-00AA0062CE6C
'ASF_Simple_Index_Object                    33000890-E5B1-11CF-89F4-00A0C90349CB
'ASF_Index_Object                           D6E229D3-35DA-11D1-9034-00A0C90349BE
'ASF_Media_Object_Index_Object              FEB103F8-12AD-4C64-840F-2A1D2F7AD48C
'ASF_Timecode_Index_Object                  3CB73FD0-0C4A-4803-953D-EDF7B6228F0C
'
'======================================================================================
'1.1    Header Object GUIDs
'The following table contains the names and values of standard ASF Header Object GUIDs.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_File_Properties_Object                 8CABDCA1-A947-11CF-8EE4-00C00C205365
'ASF_Stream_Properties_Object               B7DC0791-A9B7-11CF-8EE6-00C00C205365
'ASF_Header_Extension_Object                5FBF03B5-A92E-11CF-8EE3-00C00C205365
'ASF_Codec_List_Object                      86D15240-311D-11D0-A3A4-00A0C90348F6
'ASF_Script_Command_Object                  1EFB1A30-0B62-11D0-A39B-00A0C90348F6
'ASF_Marker_Object                          F487CD01-A951-11CF-8EE6-00C00C205365
'ASF_Bitrate_Mutual_Exclusion_Object        D6E229DC-35DA-11D1-9034-00A0C90349BE
'ASF_Error_Correction_Object                75B22635-668E-11CF-A6D9-00AA0062CE6C
'ASF_Content_Description_Object             75B22633-668E-11CF-A6D9-00AA0062CE6C
'ASF_Extended_Content_Description_Object    D2D0A440-E307-11D2-97F0-00A0C95EA850
'ASF_Content_Branding_Object                2211B3FA-BD23-11D2-B4B7-00A0C955FC6E
'ASF_Stream_Bitrate_Properties_Object       7BF875CE-468D-11D1-8D82-006097C9A2B2
'ASF_Content_Encryption_Object              2211B3FB-BD23-11D2-B4B7-00A0C955FC6E
'ASF_Extended_Content_Encryption_Object     298AE614-2622-4C17-B935-DAE07EE9289C
'ASF_Digital_Signature_Object               2211B3FC-BD23-11D2-B4B7-00A0C955FC6E
'ASF_Padding_Object                         1806D474-CADF-4509-A4BA-9AABCB96AAE8
'
'======================================================================================
'1.2    Header Extension Object GUIDs
'The following table contains the names and values of the GUIDs for the standard objects
'found inside the ASF Header Extension Object.
'--------------------------------------------------------------------------------------
'Name                                       GUID
'ASF_Extended_Stream_Properties_Object      14E6A5CB-C672-4332-8399-A96952065B5A
'ASF_Advanced_Mutual_Exclusion_Object       A08649CF-4775-4670-8A16-6E35357566CD
'ASF_Group_Mutual_Exclusion_Object          D1465A40-5A79-4338-B71B-E36B8FD6C249
'ASF_Stream_Prioritization_Object           D4FED15B-88D3-454F-81F0-ED5C45999E24
'ASF_Bandwidth_Sharing_Object               A69609E6-517B-11D2-B6AF-00C04FD908E9
'ASF_Language_List_Object                   7C4346A9-EFE0-4BFC-B229-393EDE415C85
'ASF_Metadata_Object                        C5F8CBEA-5BAF-4877-8467-AA8C44FA4CCA
'ASF_Metadata_Library_Object                44231C94-9498-49D1-A141-1D134E457054
'ASF_Index_Parameters_Object                D6E229DF-35DA-11D1-9034-00A0C90349BE
'ASF_Media_Object_Index_Parameters_Object   6B203BAD-3F11-48E4-ACA8-D7613DE2CFA7
'ASF_Timecode_Index_Parameters_Object       F55E496D-9797-4B5D-8C8B-604DFE9BFB24
'ASF_Compatibility_Object                   75B22630-668E-11CF-A6D9-00AA0062CE6C
'ASF_Advanced_Content_Encryption_Object     43058533-6981-49E6-9B74-AD12CB86D58C
'
'======================================================================================
'1.3    Stream Properties Object Stream Type GUIDs
'The following table contains the names and values of standard GUIDs for the
'Stream Type field of the Stream Properties Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Audio_Media                            F8699E40-5B4D-11CF-A8FD-00805F5C442B
'ASF_Video_Media                            BC19EFC0-5B4D-11CF-A8FD-00805F5C442B
'ASF_Command_Media                          59DACFC0-59E6-11D0-A3AC-00A0C90348F6
'ASF_JFIF_Media                             B61BE100-5B4E-11CF-A8FD-00805F5C442B
'ASF_Degradable_JPEG_Media                  35907DE0-E415-11CF-A917-00805F5C442B
'ASF_File_Transfer_Media                    91BD222C-F21C-497A-8B6D-5AA86BFC0185
'ASF_Binary_Media                           3AFB65E2-47EF-40F2-AC2C-70A90D71D343
'
'======================================================================================
'1.3.1  Web stream Type-Specific Data GUIDs
'The following table contains the names and values of the GUIDs used in the
'Type-Specific Data field of the Stream Properties Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Web_Stream_Media_Subtype               776257D4-C627-41CB-8F81-7AC7FF1C40CC
'ASF_Web_Stream_Format                      DA1E6B13-8359-4050-B398-388E965BF00C
'
'======================================================================================
'1.4    Stream Properties Object Error Correction Type GUIDs
'The following table contains the names and values of the GUIDs for the
'Error Correction Type field of the Stream Properties Object as well as
'in the Error Correction Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_No_Error_Correction                    20FB5700-5B55-11CF-A8FD-00805F5C442B
'ASF_Audio_Spread                           BFC3CD50-618F-11CF-8BB2-00AA00B4E220
'
'======================================================================================
'1.5    Header Extension Object GUIDs
'The following table contains the name and value of the GUID for the
'Reserved 1 field of the Header Extension Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Reserved_1                             ABD3D211-A9BA-11cf-8EE6-00C00C205365
'
'======================================================================================
'1.6    Advanced Content Encryption Object System ID GUIDs
'The following table contains the names and values of standard GUIDs
'for the System ID field of the Advanced Content Encryption Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Content_Encryption_System_Windows_Media_DRM_Network_Devices
'                                           7A079BB6-DAA4-4e12-A5CA-91D38DC11A8D
'
'======================================================================================
'1.7    Codec List Object GUIDs
'The following table contains the name and value of the GUIDs for the
'Reserved 2 field of the Codec List Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Reserved_2                             86D15241-311D-11D0-A3A4-00A0C90348F6
'
'======================================================================================
'1.8    Script Command Object GUIDs
'The following table contains the name and value of the GUIDs for the
'Reserved 3 field of the Script Command Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Reserved_3                             4B1ACBE3-100B-11D0-A39B-00A0C90348F6
'
'======================================================================================
'1.9    Marker Object GUIDs
'The following table contains the name and value of the GUIDs for the
'Reserved 4 field of the Marker Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Reserved_4                             4CFEDB20-75F6-11CF-9C0F-00A0C90349CB
'
'======================================================================================
'1.10   Mutual Exclusion Object Exclusion Type GUIDs
'The following table contains the names and values of the GUIDs for the
'Exclusion Type field of the Mutual Exclusion Object:
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Mutex_Language                         D6E22A00-35DA-11D1-9034-00A0C90349BE
'ASF_Mutex_Bitrate                          D6E22A01-35DA-11D1-9034-00A0C90349BE
'ASF_Mutex_Unknown                          D6E22A02-35DA-11D1-9034-00A0C90349BE
'
'======================================================================================
'1.11   Bandwidth Sharing Object GUIDs
'The following table contains the names and values of the GUIDs for the
'Sharing Type field of the Bandwidth Sharing Object.
'Name                                       GUID
'--------------------------------------------------------------------------------------
'ASF_Bandwidth_Sharing_Exclusive            AF6060AA-5197-11D2-B6AF-00C04FD908E9
'ASF_Bandwidth_Sharing_Partial              AF6060AB-5197-11D2-B6AF-00C04FD908E9
'
'======================================================================================
'1.12   Standard Payload Extension System GUIDs
'The following table contains the names and values of the GUIDs for the
'standard Payload Extension Systems.
'Name                                               GUID
'--------------------------------------------------------------------------------------
'ASF_Payload_Extension_System_Timecode'             399595EC-8667-4E2D-8FDB-98814CE76C1E
'ASF_Payload_Extension_System_File_Name             E165EC0E-19ED-45D7-B4A7-25CBD1E28E9B
'ASF_Payload_Extension_System_Content_Type          D590DC20-07BC-436C-9CF7-F3BBFBF1A4DC
'ASF_Payload_Extension_System_Pixel_Aspect_Ratio    1B1EE554-F9EA-4BC8-821A-376B74E4C4B8
'ASF_Payload_Extension_System_Sample_Duration       C6BD9450-867F-4907-83A3-C77921B733AD
'ASF_Payload_Extension_System_Encryption_Sample_ID  6698B84E-0AFA-4330-AEB2-1C0A98D7A44D
'
'======================================================================================
'Begin Declarations...
'======================================================================================
Dim WholeFile() As Byte
Dim HeaderEnd As Long
Dim Header() As Byte
Dim StreamData() As Byte
'
'Guids are most commonly written in text as a sequence of hexadecimal digits as such:
'75B22630-668E-11CF-A6D9-00AA0062CE6C
'This text notation follows from the data structure defined above. The sequence is
'    Data1 (8 characters)
'    Hyphen
'    Data2 (4 characters)
'    Hyphen
'    Data3 (4 characters)
'    Hyphen
'    Initial two items from Data4 (4 characters)
'    Hyphen
'    Remaining six items from Data4 (12 characters)
'
'main Header
Dim ASF_QSize As Currency
Dim CD_QSize As Currency
'ASF_Extended_Content_Description_Object
Dim ExtCont_QSize As Currency
'ASF_File_Properties_Object
Dim FP_QSize As Currency
Dim FileQSize As Currency
Dim calcFileQSize As Currency
Dim CreationDateQSize As Variant
Dim DataPacketsQSize As Currency
Dim PlayDurationQSize As Currency
Dim SendDurationQSize As Currency
Dim PreRollQSize As Currency
'ASF_Header_Extension_Object
Dim Ext_QSize As Currency
'Compatibility Object
Dim Comp_QSize As Currency
'ASF_Metadata_Object
Dim MetaData_QSize As Currency
Dim MetaDataLib_QSize As Currency
'ASF_Padding_Object
Dim Pad_QSize As Currency
'ASF_Extended_Stream_Properties_Object
Dim ExtStream_QSize As Currency
Dim StartTimeQSize As Currency
Dim EndTimeQSize As Currency
Dim AvgTimePerFrameQSize As Currency
Dim TimeOffsetQSize As Currency
'ASF_Codec_List_Object
Dim Codec_QSize As Currency
'ASF_Stream_Properties_Object
Dim Stream_QSize As Currency
'ASF_Stream_BitRate_Properties_Object
Dim BitRate_QSize As Currency
'ASF_Data_Object
Dim Data_QSize As Currency
Dim TotalData_QSize As Currency
Dim Packets_QSize As Currency
Dim PH_QSize As Currency
'
Dim Accum_QSize As Currency
Dim Accum_HeaderQSize As Currency
Dim AccumRead_QSize As Currency
Dim AccumWrite_QSize As Currency
'
Dim WMA_GUID As GuidStruct
Dim GUIDIn As GuidStruct
Dim GUIDOut As GuidStruct
'setup the variables for the string
Dim strNextGUID As String * 36
Dim strHeaderGUID As String * 36
Dim strCD_GUID As String * 36
Dim strInfo_GUID As String * 36
Dim strFP_GUID As String * 36
Dim strFile_GUID As String * 36
Dim strEXT_GUID As String * 36
Dim strRes1_GUID As String * 36
Dim strStream_GUID As String * 36
Dim strStreamType_GUID As String * 36
Dim strComp_GUID As String * 36
Dim strError_GUID As String * 36
Dim strMetaData_GUID As String * 36
Dim strMetaLibraryData_GUID As String * 36
Dim strPad_GUID As String * 36
Dim strExtStream_GUID As String * 36
Dim strExtSystem_GUID As String * 36
Dim strCodec_GUID As String * 36
Dim strRes2_GUID As String * 36
Dim strAudio_GUID As String * 36
Dim strBitRate_GUID As String * 36
Dim strData_GUID As String * 36
Dim strDataFileID_GUID As String * 36
Dim strPH_GUID As String * 36
'
Private Type WMAHeaderObject
    'each long is 4 bytes...our signature is 16 bytes
    'byte #------------------------------------------------|
    ' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |
    '------------|-------------|-------------|-------------|
    '30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |
    '------------|-------------|-------------|-------------|
    ObjectID As GuidStruct
    'byte #------------------------------------------>|
    '|17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30 |
    '|------------|-------------|-------------|-------|
    '|51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02 |
    '|------------|-------------|-------------|-------|
    ObjectSize As DlngSize 'Valid values are at least 30 bytes
    HeaderObjects As Long 'always 7???
    Reserved1 As Byte 'always 1
    Reserved2 As Byte 'always 2
End Type
Dim objMainHeader As WMAHeaderObject
'
'------------------------------------------------------
'Field name             Field type  Size (bytes)
'------------------------------------------------------
'Object ID              GUID        16  'ASF_Content_Description_Object
'Object Size            QWORD       8   'Valid values are at least 34 bytes.
'Title Length           WORD        2
'Author Length          WORD        2
'Copyright Length       WORD        2
'Description Length     WORD        2
'Rating Length          WORD        2
'Title                  WCHAR       Varies
'Author                 WCHAR       Varies
'Copyright              WCHAR       Varies
'Description            WCHAR       Varies
'Rating                 WCHAR       Varies
'
Private Type Content_Description_Object
    'byte #----------------------------------------------->|
    '31-------34 | 35-------38 | 39-------42 | 43-------46 |
    '33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |
    '|-----------|-------------|-------------|-------------|
    ObjectID As GuidStruct 'ASF_Content_Description_Object
    'byte #-------------------------------------------------------->|
    '|47-------50 | 51-------54 | 55-------58 | 59 60 61 62 | 63 64 |
    '|42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00 00 00 | 00 00 |
    ' ^^                          ^^    ^^                          |
    '                        TitleLen ArtistLen                     |
    '_______________________________________________________________|
    ObjectSize As DlngSize
    TitleLen As Integer
    ArtistLen As Integer
    CopyrightLen As Integer
    DescriptionLen As Integer
    RatingLen As Integer
End Type
Dim objCD_Info As Content_Description_Object
'byte 65 is the song title with a len specified by byte 55.
'immediately after that is the artist's name.
Dim aryTitle() As Byte
Dim aryArtist() As Byte
Dim aryCopyright() As Byte
Dim aryDescription() As Byte
Dim aryRating() As Byte
'
Dim wmaTitle As String
Dim wmaArtist As String
Dim wmaCopyright As String
Dim wmaDescription As String
Dim wmaRating As String
'
Dim DescriptorCount As Integer
Dim TagIndex As Integer
Dim MaxTags As Integer
Dim i As Integer
Dim j As Integer
Dim fNum As Integer
Dim eM$
Private Type Extended_Content_Description_Object
    'After the artist's name is the following
    '40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
    ObjectID As GuidStruct  'ASF_Extended_Content_Description_Object
    '
    'followed by another group that changes
    'This may be a double long to indicate the size of the variable data
    'C2 04 00 00  00 00 00 00 or
    'CE 04 00 00  00 00 00 00 or
    'DC 04 00 00  00 00 00 00
    ObjectSize As DlngSize 'Valid values are at least 26 bytes
    '
    'then i think 2 bytes indicates the number of tags
    RecordCount As Integer 'Description Records Count
End Type
Dim objExtContInfo As Extended_Content_Description_Object
'This is then followed by some indicators to determine the next data.
'integer for tagname length followed by tagname
'
'Data Type
'Specifies the type of information being stored. The following values are recognized.
'Value type  Description
'0x0000  Unicode string The data consists of a sequence of Unicode characters.
'0x0001  BYTE array.    The type of data is implementation-specific.
'0x0002  BOOL.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'                       Only 0x0000 or 0x0001 are permitted values.
'0x0003  DWORD.         The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'0x0004  QWORD.         The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'0x0005  WORD.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
Private Type DescriptionRecordEntry
    NameLen As Integer
    Name As String
    bytTagName() As Byte
    'the next 3 bytes after the tag name indicate the type of data for the tag.
    DataType  As Integer
    DataLen  As Integer
    Data As Variant
    bytTagData() As Byte
'WM/MCDI, Music CD identifier of the CD from which the file came.
End Type
Dim Descriptor() As DescriptionRecordEntry
'
'-------------------------------------------------------------------------------
Private Type FilePropertiesObject
    ObjectID As GuidStruct
    ObjectSize As DlngSize ' Valid values are at least 104 bytes
    FileID As GuidStruct
    FileSize As DlngSize 'value of this field is invalid if the Broadcast Flag is set to 1.
    CreationDate As DlngSize
    DataPackets As DlngSize
    PlayDuration As DlngSize
    SendDuration As DlngSize
    Preroll As DlngSize
    Flags As Long
    'Broadcast Flag (bit 0)
    'SeekableFlag   (bit 1)
    'Reserved             30 bits
    MinPacketSize As Long
    MaxPacketSize As Long
    MaxBitrate As Long
End Type
Dim objFileInfo As FilePropertiesObject
Dim BroadcastFlag As Byte
Dim SeekableFlag As Byte
'
'-------------------------------------------------------------------------------
Private Type HeaderExtensionObject
    ObjectID As GuidStruct 'ASF_Header_Extension_Object
    ObjectSize As DlngSize 'The value of this field shall be set to 46 bytes  ?
    Reserved1  As GuidStruct 'ASF_Reserved_1
    Reserved2 As Integer 'always 6
    ExtDataSize As Long
End Type
Dim objHeaderExtInfo As HeaderExtensionObject
Dim bytExtData() As Byte
Dim ExtDataBytesLeft As Long
Dim ExtDataBytesToWrite As Long
'
'The fields are defined as follows:
'Object ID
'Specifies the GUID for the Header Extension Object.
' The value of this field shall be set to ASF_Header_Extension_Object.
'
'Object Size
'Specifies the size, in bytes, of the Header Extension Object.
' The value of this field shall be set to 46 bytes.
'
'Reserved Field 1
'This value shall be set to ASF_Reserved_1 as defined in section 10.6.
'
'Reserved Field 2
'This value should be set to 6.
'
'-------------------------------------------------------------------------------
'Stream Properties Object
'Object ID                      GUID    16
'Object Size                    QWORD   8
'StreamType                     GUID    16
'Error Correction Type          GUID    16
'TimeOffset                     QWORD   8
'TypeDataLen                    DWORD   4
'ErrorDataLen                   DWORD   4
'Flags                          WORD    2
'   StreamNumber                  7 (LSB)
'   Reserved                      8 bits
'   Encrypted Content Flag        1 bit
'Reserved                       DWORD   4
'TypeSpecific Data              BYTE    varies
'ErrorCorrection Data           BYTE    varies
Private Type StreamPropertiesObject
    ObjectID As GuidStruct 'ASF_Stream_Properties_Object
    ObjectSize As DlngSize ' Valid values are at least 78 bytes
    StreamType As GuidStruct
    ErrorType As GuidStruct
    TimeOffset As DlngSize
    TypeDataLen As Long
    ErrorDataLen As Long
    Flags As Integer
    'StreamNumber               bits 0-6 (LSB)
    'Reserved (bits 7-14) These bits are reserved and should be set to 0.
    'Encrypted Content Flag (bit 15)
    Reserved As Long
End Type
Dim objStreamInfo As StreamPropertiesObject
Dim bytTypeSpecific() As Byte
Dim bytErrorCorrection() As Byte
'
'The fields are defined as follows:
'Object ID
'Specifies the GUID for ASF_Stream_Properties_Object.
'
'Object Size
'Specifies the size, in bytes, of the Stream Properties Object. (> 78 bytes).
'
'Stream Type
'Specifies the GUID for the type of the stream (for example, audio, video, and so on).
'
'Error Correction Type
'Specifies the error correction type used by this digital media stream.
' For streams other than audio, this value should be set to ASF_No_Error_Correction.
' For audio streams, this value should be set to ASF_Audio_Spread.
' The possible GUIDs are defined in section 10.5, and the standard error
' correction schemes are described in section 9.1.
'
'Time Offset
'Specifies the presentation time offset of the stream in 100-nanosecond units.
' The value of this field is added to all of the timestamps of the samples in the stream.
' This value shall be equal to the send time of the first interleaved packet in the data section.
' The value of this field is typically 0. It is non-zero in the case when an ASF file
' is edited and it is not possible for the editor to change the presentation times
' and send times of ASF packets. Note that if more than one stream is present in
' an ASF file the offset values of all stream properties objects must be equal.
'
'Type-Specific Data Length
'Specifies the number of bytes in the Type-Specific Data field.
'
'Error Correction Data Length
'Specifies the number of bytes in the Error Correction Data field.
'
'Flags
'The flags are stored in LSB order.
'
'Stream Number (bits 0-6)
'Specifies the number of this stream. 0 is an invalid stream.
' Valid values are between 1 and 127. The numbers assigned to streams
' in an ASF presentation may be any combination of unique values;
' parsing logic must not assume that streams are numbered sequentially.
' Reserved (bits 7-14) These bits are reserved and should be set to 0.
' Encrypted Content Flag (bit 15)
'Specifies, if set, that the data contained in this stream is encrypted
'and will be unreadable unless there is a way to decrypt the stream.
'
'Reserved
'This field is reserved and should be set to 0.
'
'Type-Specific Data
'Specifies type-specific format data.
' The structure for the Type-Specific Data field is determined by the
' value stored in the Stream Type field. The structure for the
' Type-Specific Data field for standard ASF media types is detailed in section 9.
'
'Error Correction Data
'Specifies data specific to the error correction type.
' The structure for the Error Correction Data field is determined by
' the value stored in the Error Correction Type field.
' For example, an audio data stream might need to know how codec chunks
' were redistributed, or it might need a sample of encoded silence. For
' detailed information, see section 9.1.
'-------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------
Private Type CodecListObject
    ObjectID As GuidStruct 'ASF_Codec_List_Object
    ObjectSize As DlngSize 'Valid values are at least 44 bytes.
    Reserved  As GuidStruct 'ASF_Reserved_2
    CodecEntriesCount As Long
    'Codec Entries           See below   varies
End Type
Dim objCodecListInfo As CodecListObject
Dim CodecCount As Long
'
'Field Name             Field Type  Size (bytes)
'----------------------------------------------
'Type                       WORD    2
'Codec Name Length          WORD    2
'Codec Name                 WCHAR  varies
'Codec Description Length   WORD    2
'Codec Description          WCHAR  varies
'Codec Information Length   WORD    2
'Codec Information          BYTE   varies
'
Private Type CodecRecordEntry
    Type As Integer    'get the codec type
    NameLen As Integer 'get the length of the name
    bytCodecName() As Byte 'get name data
    Name As String
    '
    DescrLen  As Integer 'Description Length
    bytCodecDescr() As Byte 'get descr data
    Descr As String
    '
    DataLen  As Integer 'Codec Data Length
    bytCodecData() As Byte 'Codec Information
    Data As Integer 'Variant
'WM/MCDI, Music CD identifier of the CD from which the file came.
End Type
Dim Codec() As CodecRecordEntry
Dim MaxCTags As Integer
Dim CTagIndex As Integer
'
'-------------------------------------------------------------------------------
'ASF_Audio_Media
'Codec ID / Format Tag      WORD        2
'Number of Channels         WORD        2
'Samples Per Second         DWORD       4
'Avg # of Bytes Per Sec     DWORD       4
'Block Alignment            WORD        2
'Bits Per Sample            WORD        2
'Codec Specific Data Size   WORD        2
'Codec Specific Data        BYTE     varies
Private Type AudioCodecInfo
    CodecID As Integer
    NumChannels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlignment As Integer
    BitsPerSample  As Integer
    CodecDataSize As Integer
End Type
Dim objAudioInfo As AudioCodecInfo
Dim bytCodecSpecificData() As Byte 'Codec Information'
'-------------------------------------------------------------------------------
'Undocumented Compatibility object...
'the guid is different than the documentation states
Private Type CompatibilityObject '(26 bytes)
    ObjectID As GuidStruct  '16 bytes
    ObjectSize As DlngSize  '8 bytes The value of this field shall be set to 26 bytes
    Profile As Byte         '1 byte 'This field is reserved and should be set to 2.
    Mode As Byte            '1 byte 'This field is reserved and should be set to 1.
End Type
Dim objCompatibilityInfo As CompatibilityObject
'
'-------------------------------------------------------------------------------
'The Metadata Object is represented using the following structure.
'
'Field name                 Field type  Size (bytes)
'Object ID                  GUID        16
'Object Size                QWORD       8
'Description Records Count  WORD        2
'Description Records        See below   varies
'
Private Type MetaDataObject '(26 bytes)
    ObjectID As GuidStruct  '16 bytes =>ASF_Metadata_Object
    ObjectSize As DlngSize  '8 bytes =>The value of this field shall be set to 26 bytes +data entry
    RecordCount As Integer ' 2 bytes =>Description Records Count
End Type
Dim objMetaDataInfo As MetaDataObject
'we also use this structure for ASF_Metadata_Library_Object
Dim objMetaDataLibraryInfo As MetaDataObject
'
'Description Records
'The structure of each Description Record entry is shown in the following table.
'Field name                 Field type  Size (bytes)
'Reserved (Must Be Zero)    WORD        2
'Stream Number              WORD        2
'Name Length                WORD        2
'Data Type                  WORD        2
'Data Length                DWORD       4
'Name                       WCHAR     varies
'Data                       See below
'
Private Type MetaDataDescriptionRecord '(12 bytes)
    Reserved As Integer 'always 0
    StreamNumber As Integer
    NameLen As Integer
    DataType As Integer
    DataLen As Long
End Type
Dim objMetaDataRecord As MetaDataDescriptionRecord
'
Private Type MetaDataDescriptionRecordEntry
    Reserved As Integer 'always 0
    StreamNumber As Integer
    NameLen As Integer
    Name As String
    bytTagName() As Byte
    'the next 3 bytes after the tag name indicate the type of data for the tag.
    DataType As Integer
    DataLen As Integer
    Data As Variant
    bytTagData() As Byte
'WM/MCDI, Music CD identifier of the CD from which the file came.
End Type
Dim MetaDescriptor() As MetaDataDescriptionRecordEntry
Dim MaxMTags As Integer
Dim MTagIndex As Integer
'
'The fields are defined as follows:
'Reserved (Must Be Zero)
' This field must contain the value 0.
'
'Stream Number
'Specifies whether the entry applies to a specific digital media stream
' or whether it applies to the whole file.
' A value of 0 in this field indicates that it applies to the whole file;
' otherwise, the entry applies only to the indicated stream number
' and must be between 1 and 127.
'
'Name Length
'Specifies the number of valid bytes stored in the Name field.
' Valid values are even numbers.
'
'Data Type
'Specifies the type of information being stored. (Standard).
'
'Data Length
'Specifies the length in bytes of the Data field.
' Valid values are less than 65536.
'
'Name
'Specifies the name that uniquely identifies the attribute being described.
' Names are case-sensitive.
'
'Data
'Specifies the actual metadata being stored.
' The Data field should be interpreted according to the value stored in the Data Type field.
'================================================================================
'
'The Padding Object is represented using the following structure.
'Field name     Field type  Size (bytes)
'------------------------------------------------
'Object ID      GUID        16
'Object Size    QWORD       8
'Padding Data   BYTE        varies 'Object Size - 24 bytes
Private Type PaddingDataObject '(24 bytes)
    ObjectID As GuidStruct  '16 bytes => ASF_Padding_Object
    ObjectSize As DlngSize  '8 bytes => Valid values are at least 24 bytes
End Type
Dim objPaddingInfo As PaddingDataObject
Dim PaddingLen As Long 'Object Size -  Len(objPaddingInfo)
Dim bytPaddingData() As Byte
'-------------------------------------------------------------------------------
'The Extended Stream Properties Object is represented using the following structure.
'
'Field name                     Field type  Size (bytes)
'-------------------------------------------------------------
'Object ID                          GUID    16
'Object Size                        QWORD   8
'Start Time                         QWORD   8
'End Time                           QWORD   8
'Data Bitrate                       DWORD   4
'Buffer Size                        DWORD   4
'Initial Buffer Fullness            DWORD   4
'Alternate Data Bitrate             DWORD   4
'Alternate Buffer Size              DWORD   4
'Alternate Initial Buffer Fullness  DWORD   4
'Maximum Object Size                DWORD   4
'Flags                              DWORD   4
'   Reliable Flag (LSB)
'   Seekable Flag
'   No Cleanpoints Flag
'   Resend Live Cleanpoints Flag
'   Reserved Flags                          28 bits
'Stream Number                      WORD        2
'Stream Language ID Index           WORD        2
'Average Time Per Frame             QWORD       8
'Stream Name Count                  WORD        2
'Payload Extension System Count     WORD        2
'Stream Names                   See below   varies
'Payload Extension Systems      See below   varies
'Stream Properties Object       See below   varies
'
Private Type ExtendedStreamPropertiesObject
    ObjectID As GuidStruct 'ASF_Extended_Stream_Properties_Object.
    ObjectSize As DlngSize
    StartTime As DlngSize
    EndTime As DlngSize
    DataBitrate As Long
    BufferSize As Long
    InitialBufferFullness As Long
    AltDataBitrate As Long
    AltBufferSize As Long
    AltInitBufferFullness As Long
    MaxObjectSize As Long
    Flags As Long
    '   Reliable Flag (LSB)
    '   Seekable Flag
    '   No Cleanpoints Flag
    '   Resend Live Cleanpoints Flag
    '   Reserved Flags                          28 bits
    StreamNumber As Integer
    StreamLangIDIndex As Integer
    AvgTimePerFrame As DlngSize
    StreamNameCount As Integer
    PayloadExtSystemCount As Integer
End Type
Dim objExtStreamInfo As ExtendedStreamPropertiesObject
'Stream Names
'The structure of each Stream Name entry is shown in the following table.
'Field name             Field type  Size (bytes)
'-----------------------------------------------------------------------------
'Language ID Index      WORD        2
'Stream Name Length     WORD        2
'Stream Name            WCHAR   varies
Dim LangIDIndex As Integer
Dim StreamNLen As Integer
Dim bytStreamName() As Byte
'below array stores the name info
Dim StreamName() As String
'
'Payload extension systems are described in the following table.
'Field name  Field type  Size (bits)
''-----------------------------------------------------------------------------
'Extension System ID             GUID    16
'Extension Data Size             WORD    2
'Extension System Info Length    DWORD   4
'Extension System Info           BYTE    varies
Private Type PayloadExtensionSystems
    ExtSystemID As GuidStruct
    ExtDataSize As Integer
    ExtSystemInfoLen As Long
End Type
Dim objPayloadInfo As PayloadExtensionSystems
'below gets the bytes for each ExtSystemInfo
Dim bytExtSystemInfo() As Byte
'below is an array for each ExtSystemInfo
Dim ExtSystemInfo() As Variant
'
'Stream Bitrate Properties Object (optional but recommended, one only)
'The Stream Bitrate Properties Object defines the average bit rate
' of each digital media stream.
' It is represented using the following structure.
'
'Field name             Field type  Size (bytes)
'---------------------------------------------------------
'Object ID              GUID        16
'Object Size            QWORD       8
'Bitrate Records Count  WORD        2
'Bitrate Records        See below   varies
'
Private Type StreamBitrateProperties
    ObjectID As GuidStruct 'ASF_Stream_Bitrate_Properties_Object
    ObjectSize As DlngSize 'at least 26 bytes
    BitrateRecords As Integer
End Type
Dim objBitRateInfo As StreamBitrateProperties
'
Private Type BitrateRecordEntry
    Flags As Integer
    AvgBitrate As Long
End Type
Dim BitRateEntry() As BitrateRecordEntry
'Bitrate Records
'The structure of each Bitrate Record entry is shown in the following table.
'Field name         Field type  Size (bits)
'------------------------------------------------------
'Flags              WORD    2
'  -Stream Number           -7 bits (LSB)
'  -Reserved                -9 bits
'Average Bitrate    DWORD   4
'
'================================================================
'Data object
'Field name             Field type  Size (bytes)
'---------------------------------------------------------
'Object ID              GUID        16' ASF_Data_Object
'Object Size            QWORD       8 'at least 50 bytes
'File ID                GUID        16
'Total Data Packets     QWORD       8
'Reserved               WORD        2 'set to 0x0101. (257D)
Private Type DataObject
    ObjectID As GuidStruct 'ASF_Stream_Bitrate_Properties_Object
    ObjectSize As DlngSize 'at least 50 bytes
    FileID As GuidStruct
    TotalDataPackets As DlngSize 'must be equal to the Data Packet Count field in the File Properties Object
    Reserved As Integer 'The value of this field shall set to 0x0101. (257D)
End Type
Dim objDataInfo As DataObject
'
Private Type PlaceholderObject
    ObjectID As GuidStruct 'ASF_Index_Placeholder_Object
    ObjectSize As DlngSize
    Word1 As Integer
    Word2 As Integer
    Word3 As Integer
    Word4 As Integer
    Word5 As Integer
End Type
Dim objPlaceholderInfo As PlaceholderObject
'
'================================================================================
Const EmptyGuid = "00000000-0000-0000-0000-000000000000"
Const ASF_Header_Object = "75B22630-668E-11CF-A6D9-00AA0062CE6C"
Const ASF_Content_Description_Object = "75B22633-668E-11CF-A6D9-00AA0062CE6C"
Const ASF_Extended_Content_Description_Object = "D2D0A440-E307-11D2-97F0-00A0C95EA850"
Const ASF_File_Properties_Object = "8CABDCA1-A947-11CF-8EE4-00C00C205365"
Const ASF_Header_Extension_Object = "5FBF03B5-A92E-11CF-8EE3-00C00C205365"
Const ASF_Reserved_1 = "ABD3D211-A9BA-11CF-8EE6-00C00C205365"
Const ASF_Reserved_2 = "86D15241-311D-11D0-A3A4-00A0C90348F6"
Const ASF_Codec_List_Object = "86D15240-311D-11D0-A3A4-00A0C90348F6"
Const ASF_Audio_Media = "F8699E40-5B4D-11CF-A8FD-00805F5C442B"
Const ASF_Stream_Properties_Object = "B7DC0791-A9B7-11CF-8EE6-00C00C205365"
Const ASF_Audio_Spread = "BFC3CD50-618F-11CF-8BB2-00AA00B4E220"
Const ASF_Stream_Bitrate_Properties_Object = "7BF875CE-468D-11D1-8D82-006097C9A2B2"
Const ASF_Extended_Stream_Properties_Object = "14E6A5CB-C672-4332-8399-A96952065B5A"
Const ASF_Metadata_Object = "C5F8CBEA-5BAF-4877-8467-AA8C44FA4CCA"
Const ASF_Padding_Object = "1806D474-CADF-4509-A4BA-9AABCB96AAE8"
Const ASF_Undocumented_Compatibility_GUID = "26F18B5D-4584-47EC-9F5F-0E651F0452C9"
'above referred to as ASF_Compatibility_Object, but the guid is different
Const ASF_Index_Placeholder_Object = "D9AADE20-7C17-4F9C-BC28-8555DD98E2A2"
Const ASF_Data_Object = "75B22636-668E-11CF-A6D9-00AA0062CE6C"
Const ASF_Metadata_Library_Object = "44231C94-9498-49D1-A141-1D134E457054"
'below is the header for temp file created after streaming
Const ASF_TempFileHeader = "75B22668-668E-11CF-A6D9-00AA0062CE6C"
'================================================================================
'variables for the GUID
Dim NextGUID As GuidStruct
Dim bytGUID() As Byte
'
Dim tempNode As Node
Dim TV As TreeView
Dim LastAddress$
Dim AudioFileName As String
Dim DebugMode As Byte
Dim StartAddr As Long
Dim HeaderCompleteCount As Integer
Dim SubHeaderCompleteCount As Integer
Dim TempString() As String
Dim VarOutput() As Variant
Dim ExtContSize As Long
Dim PaddingEndPos As Long
Dim TypeSpecificData As String
Dim ErrorCorrectionData As DlngSize
Dim ErrorCorrectionDataQSize As Currency
Dim EndOfDataAddress As Long
Dim ObjectStack() As String
Dim StackNum As Integer
Dim FileInfoText() As String
Dim FileCount As Integer
Dim IsInIDE As Boolean
Dim lngExcessData As Long

Sub ReadWMAHeader(fName As String)
On Error GoTo InIDE
DebugMode = 0
IsInIDE = False
Debug.Print 1 / 0
GoTo NoIDE
InIDE:
IsInIDE = True
NoIDE:
On Error GoTo Oops
'
'fNum = FreeFile
'Open FName For Binary As fNum
'ReDim WholeFile(LOF(fNum))
'Get #fNum, , WholeFile 'this gets the whole dang thing
'Close #fNum
'
'byte #------------------------------------------------------------------------------------------------->
' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30 |
'------------|-------------|-------------|-------------|------------|-------------|-------------|-------|
'30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02 |
'            |             |             |             |^^          |             |             |       |------------>|
'31-------34 | 35-------38 | 39-------42 | 43-------46 |47-------50 | 51-------54 | 55-------58 | 59 60 61 62 | 63 64 |
'33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00 00 00 | 00 00 |
'                                                       ^^                          ^^    ^^                          |
'                                                                               TitleLen ArtistLen                    |
'_____________________________________________________________________________________________________________________|
'byte 65 is the song title with a len specified by byte 55.
'immediately after that is the artist's name.
'After the artist's name is the following
'40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
'followed by another group that changes
'C2 04 00 00  00 00 00 00 or
'CE 04 00 00  00 00 00 00 or
'DC 04 00 00  00 00 00 00
'
'then i think 2 bytes indicates the length of tagname and the tagname data
'This is then followed by length of the data for that tag and then the data.
'14 00   precedes WM/Lyrics (20 chars=14h)
'12 00   precedes WM/Track (18 chars=12h)
'00 10 00 is before WM/Year....
'20 00 Encoding
'I guess each tag is a given lookup and we read the lookup #; the tag name; and then
'the length and data.
'the entire header is up to 1300H 4864D
'it also appears that the tags are not always in the same order.
'reset our tag info
Set TV = frmMain.tvHeader
TV.Nodes.Clear
'add root node
Set tempNode = TV.Nodes.Add(, , "Root", fName, 1)
TV.Nodes(1).Expanded = True
'
Dim myFSO As New FileSystemObject
Dim myTextStream As TextStream
'
AudioFileName = fName
'we only need to do this to show the header
'Set myTextStream = myFSO.OpenTextFile(fName, ForReading, False)
'fstrFullHeader = myTextStream.Read(Val("&H3000"))
'myTextStream.Close
'
TagIndex = 0
MaxTags = 0
ReDim Descriptor(MaxTags)
'How do i clear the contents of a udt????
'below works, but i need to do this for each???
With objMainHeader
    .HeaderObjects = 0
    .ObjectID = GUIDStrToGUIDStruct(EmptyGuid)
    .ObjectSize.High = 0
    .ObjectSize.Low = 0
    .Reserved1 = 0
    .Reserved2 = 0
End With
'below is not valid.  Is there something similar for UDT's????
'Set objMainHeader = Nothing
'Set objCD_Info = Nothing
'Set objExtContInfo = Nothing
'Set objFileInfo = Nothing
'Set objHeaderExtInfo = Nothing
'Set objCompatibilityInfo = Nothing
'Set objMetaDataInfo = Nothing
'Set objPaddingInfo = Nothing
'Set objExtStreamInfo = Nothing
'Set objPlaceholderInfo = Nothing
'Set objMetaDataLibraryInfo = Nothing
'Set objCodecListInfo = Nothing
'Set objStreamInfo = Nothing
'Set objBitRateInfo = Nothing
'Set objDataInfo = Nothing
'
'Set objMetaDataRecord = Nothing
'Set objPayloadInfo = Nothing
'_____________________________________________________________________________________________________________________|
'beginning of parsing
fNum = FreeFile
LastAddress$ = "00000H"
Open fName For Binary As fNum
InfoTag.FileSize = LOF(fNum)
frmMain.txtStatus.SelText = "====================================================================" & vbCrLf
frmMain.txtStatus.SelText = "Reading Tag Data from " & fName & vbCrLf
eM$ = "File Length FLen = " & Format(FileLen(fName), "#,###") & " (" & Hex$(FileLen(fName)) & ")" & vbCrLf
eM$ = eM$ & "File Length LOF  = " & Format(LOF(fNum), "#,###") & " (" & Hex$(LOF(fNum)) & ")" & vbCrLf
Debug.Print eM$
eM$ = ""
fstrInfo = Format(LOF(fNum), "#,###") & " (" & Hex$(LOF(fNum)) & ")" & " | "
Dim WatchDog#
WatchDog# = Timer
HeaderCompleteCount = 0
SubHeaderCompleteCount = 0
StackNum = 0
ReDim ObjectStack(StackNum + 1)
Do
    TestNextGUID
    ReDim Preserve ObjectStack(StackNum + 1)
    ObjectStack(StackNum) = ReturnGUIDName(strNextGUID)
    Select Case strNextGUID
        Case Is = ASF_TempFileHeader, ASF_Header_Object
            ParseMainHeader
        Case Is = ASF_Content_Description_Object
            ParseContentHeader
        Case Is = ASF_Extended_Content_Description_Object
            ParseExtContent
        Case Is = ASF_File_Properties_Object
            ParseFileProperties
        '
        Case Is = ASF_Header_Extension_Object
            ParseHeaderExtension
        '
        '--------------------------------------------------
        'Below data is all part of the extended header
        'they don't count as header objects
        '
        Case Is = ASF_Undocumented_Compatibility_GUID
            ParseCompatibility
        Case Is = ASF_Metadata_Object
            ParseMetaData
        Case Is = ASF_Padding_Object
            ParsePadding
        Case Is = ASF_Extended_Stream_Properties_Object
            ParseExtendedStream
        Case Is = ASF_Index_Placeholder_Object
            ParsePlaceholder
        Case Is = ASF_Metadata_Library_Object
            ParseMetadataLibrary
        '--------------------------------------------------
        Case Is = ASF_Codec_List_Object
            ParseCodecList
        Case Is = ASF_Stream_Properties_Object
            ParseStreamProperties
        Case Is = ASF_Stream_Bitrate_Properties_Object
            ParseStreamBitrate
        Case Is = ASF_Data_Object
            ParseData
            'Next comes the data packets
            'An ASF Data Packet is a conveniently sized grouping of complete or
            'fragmented media objects from several digital media streams.
            '
        Case Else
            eM$ = "Unhandled GUID!" & vbCrLf
            eM$ = eM$ & ReturnGUIDName(strNextGUID) & vbCrLf
            eM$ = eM$ & strNextGUID & vbCrLf
            eM$ = eM$ & "at " & Seek(fNum) & " (" & AddZero(Hex$(Seek(fNum)), 6) & "H)"
            MsgBox eM$, vbCritical, "Unhandled GUID!"
            eM$ = ""
            WatchDog# = Timer
            Exit Do
    End Select
    If IsInIDE = False Then
        If Timer - WatchDog# > 60 Then
            eM$ = "WatchDog TimeOut!" & vbCrLf
            MsgBox eM$, vbCritical, "Error in ReadWMAHeader"
            eM$ = ""
            Exit Do
        End If
    End If
    StackNum = StackNum + 1
Loop Until HeaderCompleteCount >= objMainHeader.HeaderObjects
Dim ObjNum As Integer
Debug.Print "----------------------------File Structure----------------------------"
For ObjNum = 0 To StackNum - 1
    Debug.Print ObjNum; " => "; ObjectStack(ObjNum)
Next ObjNum
Debug.Print "----------------------------------------------------------------------"
'--------------------------------------------------------------------------------
AccumRead_QSize = Accum_QSize
'
Debug.Print "ASF_QSize reported = "; ASF_QSize
Debug.Print "Accum_HeaderQSize  = "; Accum_HeaderQSize
If Accum_HeaderQSize <> ASF_QSize Then
    eM$ = "Accum_HeaderQSize <> ASF_QSize"
    eM$ = eM$ & vbCrLf & "Accum_HeaderQSize  = " & Accum_HeaderQSize
    eM$ = eM$ & vbCrLf & "Main Header ASF_QSize = " & ASF_QSize
    eM$ = eM$ & vbCrLf & "Difference               = " & ASF_QSize - Accum_HeaderQSize
    eM$ = eM$ & vbCrLf
    eM$ = eM$ & vbCrLf & "Do you want to use the Accum_HeaderQSize as the Header Object size?"
    eTitle$ = App.Title & ": Accumulated Size Error"
    mError = MsgBox(eM$, vbYesNo, eTitle$)
    If mError = vbYes Then
        ASF_QSize = Accum_HeaderQSize
    End If
    eM$ = ""
End If
'
'lets check the added size of the file properties object for fun
Debug.Print "FileQSize reported = "; FileQSize
'Calculate the size for use in the File Properties
calcFileQSize = Len(objMainHeader)
calcFileQSize = calcFileQSize + QWORD(objCD_Info.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objExtContInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objFileInfo.ObjectSize)
'don't add the Header Extension object size to the accum. value as it
'encompasses several other objects that are already counted.
'Add the length of the object instead.
'calcFileQSize = calcFileQSize + QWORD(objHeaderExtInfo.ObjectSize)
calcFileQSize = calcFileQSize + Len(objHeaderExtInfo)
calcFileQSize = calcFileQSize + QWORD(objCompatibilityInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objMetaDataInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objPaddingInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objExtStreamInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objPlaceholderInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objMetaDataLibraryInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objCodecListInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objStreamInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objBitRateInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objDataInfo.ObjectSize)
Debug.Print "calcFileQSize Calculated = "; calcFileQSize
'
'
SanityCheck:
Debug.Print "-------------Below 4 should be the same---------------"
Debug.Print LOF(fNum), AddZero(Hex$(LOF(fNum)), 6); "H is the Length FileSize"
Debug.Print FileQSize, AddZero(Hex$(FileQSize), 6); "H is the FileQSize from File Properties"
Debug.Print calcFileQSize, AddZero(Hex$(calcFileQSize), 6); "H is the calculated File Size for File Properties"
Debug.Print Accum_QSize, AddZero(Hex$(Accum_QSize), 6); "H is the Accum_QSize"
Debug.Print EndOfDataAddress, AddZero(Hex$((EndOfDataAddress)), 6); "H is LastAddress"
'
Debug.Print Accum_HeaderQSize, AddZero(Hex$(Accum_HeaderQSize), 6); "H is the Accum_HeaderQSize"
Debug.Print ASF_QSize, AddZero(Hex$(ASF_QSize), 6); "H is the Header length"
Debug.Print "-------------Below should be the same---------------"
Debug.Print LOF(fNum) - ASF_QSize, AddZero(Hex$(LOF(fNum) - ASF_QSize), 6); " is the Size - Header length"
Debug.Print Data_QSize, AddZero(Hex$(Data_QSize), 6); " is the size of the data"
Debug.Print "Excess Data Count = "; vbTab; lngExcessData; " ("; Hex$(lngExcessData); ")"
'
If (FileQSize = Accum_QSize) And (Accum_QSize = LOF(fNum)) And (LOF(fNum) = (EndOfDataAddress)) And (calcFileQSize = FileQSize) Then
    'everything matches...yay
    Debug.Print "All sizes match..."
    frmMain.txtStatus.SelText = "All sizes match..." & vbCrLf
Else
    eM$ = eM$ & vbCrLf & vbCrLf & "*** Error in FileSize *** "
    eM$ = eM$ & vbCrLf & "Actual FileSize           = " & vbTab & InfoTag.FileSize & " (" & Hex$(InfoTag.FileSize) & ")"
    eM$ = eM$ & vbCrLf & "FP Object FileSize      = " & vbTab & FileQSize & " (" & Hex$(FileQSize) & ")"
    If (FileQSize <> InfoTag.FileSize) Then
        eM$ = eM$ & vbCrLf & "Difference                = " & vbTab & InfoTag.FileSize - FileQSize & " (" & Hex$(InfoTag.FileSize - FileQSize) & ")"
    End If
    If (calcFileQSize <> FileQSize) And BroadcastFlag = 0 Then
        eM$ = eM$ & vbCrLf & vbCrLf & "*** Error in File Properties FileSize ***"
        eM$ = eM$ & vbCrLf & "Calculated FileSize      = " & vbTab & calcFileQSize & " (" & Hex$(calcFileQSize) & ")"
        eM$ = eM$ & vbCrLf & "Difference                = " & vbTab & FileQSize - calcFileQSize & " (" & Hex$(FileQSize - calcFileQSize) & ")"
        eM$ = eM$ & vbCrLf & "--------------------------------------------------------------"
    Else
        eM$ = eM$ & vbCrLf & "Calculated FileSize      = " & vbTab & calcFileQSize & " (" & Hex$(calcFileQSize) & ")"
    End If
    eM$ = eM$ & vbCrLf & "Accum. FileSize          = " & vbTab & Accum_QSize & " (" & Hex$(Accum_QSize) & ")"
    eM$ = eM$ & vbCrLf & "Last Read address      = " & vbTab & EndOfDataAddress & " (" & Hex$(EndOfDataAddress) & ")"
    eM$ = eM$ & vbCrLf & "Excess Data Count     = " & vbTab & lngExcessData & " (" & Hex$(lngExcessData) & ")"
    eM$ = eM$ & vbCrLf & "FileQSize-LastAddress  = " & vbTab & FileQSize - EndOfDataAddress & " (" & Hex$(FileQSize - EndOfDataAddress) & ")"
    frmMain.txtStatus.SelText = eM$ & vbCrLf
    'eM$ = ""
    If lngExcessData > 0 Then
        TestNextGUID
    End If
End If
If Len(eM$) > 0 Then
    MsgBox eM$, , "ReadWMAHeader"
    frmMain.txtStatus.SelText = eM$ & vbCrLf
    eM$ = ""
End If
'
Dim bytExcessData() As Byte
Dim strExcessData$
'if we have any data left, let's read it and show it in hex format
If AccumRead_QSize <> LOF(fNum) Then
    Debug.Print "AccumRead_QSize = "; AccumRead_QSize
    Debug.Print "File Length = "; LOF(fNum)
    TestNextGUID
    'Clipboard.Clear
    'Clipboard.SetText strNextGUID
    frmMain.txtStatus.SelText = "--------------------------------------------------------------------------" & vbCrLf
    frmMain.txtStatus.SelText = "strNextGUID = " & strNextGUID & vbCrLf
    'We have more extended data that's not handled...."
    Debug.Print "Hex Dump starting at "; AccumRead_QSize; " ("; Hex$(AccumRead_QSize); "H)"
    ReDim bytExcessData(lngExcessData)
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "ExcessData", LastAddress$ & ": " & "Excess Data", 3)
    Get #fNum, , bytExcessData
    For i = 0 To UBound(bytExcessData) - 1
        If i Mod 16 = 0 Then
            Debug.Print strExcessData$
            Set tempNode = TV.Nodes.Add("ExcessData", tvwChild, "strExcessData" & i, strExcessData$, 4)
            'Debug.Print AddZero(Hex$((i)), 6); vbTab;
            strExcessData$ = AddZero(Hex$((i)), 6) & "  "
        End If
        strExcessData$ = strExcessData$ & AddZero(Hex$(bytExcessData(i)), 2) & " "
    Next i
    Debug.Print
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print Seek(fNum); " ("; LastAddress$; ") After bytExcessData"
    Debug.Print "Excess Data Count = "; vbTab; lngExcessData; " ("; Hex$(lngExcessData); ")"
End If
GoTo AllDone
'===============================================================================
AllDone:
Close #fNum
If Len(eM$) > 0 Then
    MsgBox eM$, , "End of ReadWMAHeader"
    eM$ = ""
End If
'
FileCount = FileCount + 1
ReDim Preserve FileInfoText(FileCount)
Dim aryLines() As String
ReDim aryLines(1 To TV.Nodes.Count)
For i = 1 To TV.Nodes.Count
    'Debug.Print TV.Nodes(i).Text
    aryLines(i) = TV.Nodes(i).Text
Next i
FileInfoText(FileCount) = Join(aryLines, vbCrLf)
'frmMain.txtInfo.SelText = FileInfoText(FileCount)
GoTo Exit_ReadWMAHeader
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ReadWMAHeader "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ReadWMAHeader"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ReadWMAHeader:
End Sub

Sub ParseMainHeader()
On Error GoTo Oops
'------------------------------ASF_Header_Object----------------------------------------
'get the beginning main header
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ASF_Header_Object"
Get #fNum, , objMainHeader  'this gets the first block of 30 bytes of the first 64
Debug.Print "*****************************************************************************"
Debug.Print "--------------------------ASF_Header_Object----------------------------------"
With objMainHeader
    GUIDIn = .ObjectID
    strHeaderGUID = GUIDToString(GUIDIn)
    eM$ = "Executing " & ReturnGUIDName(strHeaderGUID) & vbCrLf & strHeaderGUID
    Debug.Print eM$
    If DebugMode = 1 Then
        MsgBox eM$
    End If
    eM$ = ""
    If strHeaderGUID <> ASF_Header_Object Then
        If strHeaderGUID = ASF_TempFileHeader Then
            'we have a temp file
        Else
            eM$ = "objMainHeader"
            eM$ = eM$ & vbCrLf & "---------------------------"
            eM$ = eM$ & vbCrLf & "** Error in strHeaderGUID = " & strHeaderGUID
        End If
    End If
    ASF_QSize = QWORD(.ObjectSize) 'Valid values are at least 30 bytes
    'the object size for the main header is the whole length of the header object
    Accum_HeaderQSize = Len(objMainHeader)  ' 30 ' ASF_QSize
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    fstrInfo = fstrInfo & AddSpace(ASF_QSize & " (" & Hex$(ASF_QSize) & ")", 10) & " | "
    Debug.Print "ObjectSize = "; ASF_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
'
'byte #------------------------------------------------------------------------------------------------->
' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30 |
'------------|-------------|-------------|-------------|------------|-------------|-------------|-------|
'30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02 |
'            |             |             |             |^^          |             |             |       |------------>|
    If .HeaderObjects <> 7 Then 'always 7?
        eM$ = eM$ & vbCrLf & "** Error in HeaderObjects = " & .HeaderObjects & " (" & Hex$(.HeaderObjects) & ")"
    Else
        Debug.Print "HeaderObjects OK   = "; .HeaderObjects; " - always 7"
    End If
    If .Reserved1 <> 1 Then 'always 1
        eM$ = eM$ & vbCrLf & "** Error in Reserved1 = " & .Reserved1 & " (" & Hex$(.Reserved1) & ")"
    Else
        Debug.Print "Reserved1 OK       = "; .Reserved1; " - always 1"
    End If
    If .Reserved2 <> 2 Then 'always 2
        eM$ = eM$ & vbCrLf & "** Error in Reserved2 = " & .Reserved2 & " (" & Hex$(.Reserved2) & ")"
    Else
        Debug.Print "Reserved2 OK       = "; .Reserved2; " - always 2"
    End If
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Main", LastAddress$ & ": " & "ASF_Header_Object", 3)
    Set tempNode = TV.Nodes.Add("Main", tvwChild, "GUID", "GUID = " & strHeaderGUID, 4)
    Set tempNode = TV.Nodes.Add("Main", tvwChild, "Size", "ObjectSize    = " & ASF_QSize & " (" & Hex$(ASF_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Main", tvwChild, "HeaderObjects", "HeaderObjects = " & .HeaderObjects, 4)
    Set tempNode = TV.Nodes.Add("Main", tvwChild, "Reserved1", "Reserved1     = " & .Reserved1, 4)
    Set tempNode = TV.Nodes.Add("Main", tvwChild, "Reserved2", "Reserved2     = " & .Reserved2, 4)
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Header_Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ASF_Header_Object"
'
GoTo Exit_ParseMainHeader
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseMainHeader "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseMainHeader"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseMainHeader:
End Sub

Sub ParseContentHeader()
On Error GoTo Oops
Debug.Print "--------------------------ASF_Content_Description_Object----------------------------------"
'now get the content description
'get the beginning main header
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ASF_Content_Description_Object"
Get #fNum, , objCD_Info  'this gets the next block of 64 bytes
HeaderCompleteCount = HeaderCompleteCount + 1
Debug.Print Len(objCD_Info)
With objCD_Info
    GUIDIn = .ObjectID
    strCD_GUID = GUIDToString(GUIDIn)
    eM$ = "Executing " & ReturnGUIDName(strCD_GUID) & vbCrLf & strCD_GUID
    Debug.Print eM$
    If DebugMode = 1 Then
        MsgBox eM$
    End If
    eM$ = ""
    '
    If strCD_GUID <> ASF_Content_Description_Object Then
        eM$ = "ASF_Content_Description_Object"
        eM$ = eM$ & vbCrLf & "---------------------------"
        eM$ = eM$ & vbCrLf & "** Error in strCD_GUID = " & strCD_GUID
    End If
    CD_QSize = QWORD(.ObjectSize) '34 bytes + size of content
    Accum_HeaderQSize = Accum_HeaderQSize + CD_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseContentHeader"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    fstrInfo = fstrInfo & AddSpace(CD_QSize & " (" & Hex$(CD_QSize) & ")", 10) & " | "
    Debug.Print fstrInfo
    '
    'now set the len for the title tag
    If .TitleLen > 0 Then
        ReDim aryTitle(.TitleLen - 1)
        Get #fNum, , aryTitle 'this gets the block of bytes for the title
        'wmaTitle = Replace(StrConv(aryTitle, vbUnicode), Chr$(0), "")
        wmaTitle = aryTitle
        wmaTitle = Left$(wmaTitle, Len(wmaTitle) - 1)
        InfoTag.Title = wmaTitle
    End If
    '
    'now set the len for the artist tag
    If .ArtistLen > 0 Then
        ReDim aryArtist(.ArtistLen - 1)
        Get #fNum, , aryArtist 'this gets the block of bytes for the Artist
        'wmaArtist = Replace(StrConv(aryArtist, vbUnicode), Chr$(0), "")
        wmaArtist = aryArtist
        wmaArtist = Left$(wmaArtist, Len(wmaArtist) - 1)
        InfoTag.Artist = wmaArtist
    End If
    '
    'now set the len for the Copyright tag
    If .CopyrightLen > 0 Then
        ReDim aryCopyright(.CopyrightLen - 1)
        Get #fNum, , aryCopyright 'this gets the block of bytes for the Copyright
        'wmaCopyright = Replace(StrConv(aryCopyright, vbUnicode), Chr$(0), "")
        wmaCopyright = aryCopyright
        wmaCopyright = Left$(wmaCopyright, Len(wmaCopyright) - 1)
        InfoTag.Copyright = wmaCopyright
        InfoTag.IsCopyright = True
    Else
        InfoTag.IsCopyright = False
    End If
    '
    'now set the len for the Description tag
    If .DescriptionLen > 0 Then
        ReDim aryDescription(.DescriptionLen - 1)
        Get #fNum, , aryDescription 'this gets the block of bytes for the Description
        'wmaDescription = Replace(StrConv(aryDescription, vbUnicode), Chr$(0), "")
        wmaDescription = aryDescription
        wmaDescription = Left$(wmaDescription, Len(wmaDescription) - 1)
        InfoTag.Description = wmaDescription
    End If
    '
    'now set the len for the Rating tag
    If .RatingLen > 0 Then
        ReDim aryRating(.RatingLen - 1)
        Get #fNum, , aryRating 'this gets the block of bytes for the Rating
        'wmaRating = Replace(StrConv(aryRating, vbUnicode), Chr$(0), "")
        wmaRating = aryRating
        wmaRating = Left$(wmaRating, Len(wmaRating) - 1)
        InfoTag.Rating = wmaRating
    End If
'----------------------------------------------------------------------
'
'31-------34 | 35-------38 | 39-------42 | 43-------46 |47-------50 | 51-------54 | 55-------58 | 59 60 61 62 | 63 64 |
'33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00 00 00 | 00 00 |
'                                                       ^^                          ^^    ^^                          |
'
    Debug.Print "CD_Size        = "; CD_QSize; "("; Hex$(CD_QSize); ")"
    Debug.Print "TitleLen       = "; .TitleLen; "("; Hex$(.TitleLen); ")" ' variable
    Debug.Print "Title          = "; wmaTitle
    Debug.Print "ArtistLen      = "; .ArtistLen; "("; Hex$(.ArtistLen); ")"             ' variable
    Debug.Print "Artist         = "; wmaArtist
    Debug.Print "CopyrightLen   = "; .CopyrightLen; "("; Hex$(.CopyrightLen); ")"              ' variable
    Debug.Print "Copyright      = "; wmaCopyright
    Debug.Print "DescriptionLen = "; .DescriptionLen; "("; Hex$(.DescriptionLen); ")"               ' variable
    Debug.Print "Description    = "; wmaDescription
    Debug.Print "RatingLen      = "; .RatingLen; "("; Hex$(.RatingLen); ")"              ' variable
    Debug.Print "Rating         = "; wmaRating
    '
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Content", LastAddress$ & ": " & "ASF_Content_Description_Object", 3)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "CD_GUID", "GUID = " & strCD_GUID, 4)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "CD_Size", "CD_Size      = " & CD_QSize & " (" & Hex$(CD_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "Title", "Title        = " & wmaTitle, 4)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "Artist", "Artist       = " & wmaArtist, 4)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "Copyright", "Copyright    = " & wmaCopyright, 4)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "Description", "Description  = " & wmaDescription, 4)
    Set tempNode = TV.Nodes.Add("Content", tvwChild, "Rating", "Rating       = " & wmaRating, 4)
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Content_Description_Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ASF_Content_Description_Object"
'
'ParseExtContent
GoTo Exit_ParseContentHeader
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseContentHeader "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseContentHeader"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseContentHeader:
End Sub

Sub ParseExtContent()
On Error GoTo Oops
'------------------------------ASF_Extended_Content_Description_Object---------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ParseExtContent"
Get #fNum, , objExtContInfo 'this gets a block of 26 bytes after the title/artist info
HeaderCompleteCount = HeaderCompleteCount + 1
'
eM$ = ""
Set tempNode = TV.Nodes.Add("Root", tvwChild, "ExtCont", LastAddress$ & ": " & "ASF_Extended_Content_Description_Object", 3)
With objExtContInfo
    'After the artist's name is the following
    '40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
    'which is the ASF_Extended_Content_Description_Object
    GUIDIn = .ObjectID
    strInfo_GUID = GUIDToString(GUIDIn)
    Debug.Print "strInfo_GUID= "; strInfo_GUID
    eM$ = "Executing " & ReturnGUIDName(strInfo_GUID) & vbCrLf & strInfo_GUID
    Debug.Print eM$
    If DebugMode = 1 Then
        MsgBox eM$
    End If
    eM$ = ""
    If strInfo_GUID <> ASF_Extended_Content_Description_Object Then
        eM$ = eM$ & vbCrLf & "___________________________"
        eM$ = eM$ & vbCrLf & vbCrLf & "objExtContInfo"
        eM$ = eM$ & vbCrLf & "---------------------------"
        eM$ = eM$ & vbCrLf & "** Error in strInfo_GUID = " & strInfo_GUID
    End If
    '
    'followed by another group that changes
    'C2 04 00 00  00 00 00 00 or
    ExtCont_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + ExtCont_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseExtContent"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    'Valid values are at least 26 bytes
    Debug.Print "ObjectSize       = "; ExtCont_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    fstrInfo = fstrInfo & AddSpace(ExtCont_QSize & " (" & Hex$(ExtCont_QSize) & ")", 10) & " | "
    '
    'then i think 2 bytes indicates the number of tags
    'This is then followed by some indicators to determine the next data.
'    eM$ = eM$ & vbCrLf & "RecordCount = " & .RecordCount & " (" & Hex$(.RecordCount) & ")"
    Debug.Print "RecordCount      = "; .RecordCount; "("; Hex$(.RecordCount); ")"
    DescriptorCount = .RecordCount
    Set tempNode = TV.Nodes.Add("ExtCont", tvwChild, "EC_GUID", "GUID = " & strInfo_GUID, 4)
    Set tempNode = TV.Nodes.Add("ExtCont", tvwChild, "EC_Size", "ObjectSize    = " & ExtCont_QSize & " (" & Hex$(ExtCont_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtCont", tvwChild, "EC_RecordCount", "RecordCount   = " & .RecordCount, 4)
End With
'Debug.Print Len(objExtContInfo)
Debug.Print fstrInfo
Debug.Print eM$
If Len(eM$) > 0 Then
    MsgBox eM$, , "Extended Content Info " '& wmaTitle
    eM$ = ""
End If
'
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ParseExtContent"
ParseExtContentData
GoTo Exit_ParseExtContent
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseExtContent "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseExtContent"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseExtContent:
End Sub

Sub ParseExtContentData()
On Error GoTo Oops
Dim strInfo$
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ParseExtContentData"
ReDim Preserve Descriptor(MaxTags + DescriptorCount)
eM$ = ""
'Content Descriptors
'The structure of each Content Descriptor entry is shown in the following table.
'Field Name                 Field Type  Size (bytes)
'Descriptor Name Length         WORD    2
'Descriptor Name                WCHAR   varies
'Descriptor Value Data Type     WORD    2
'Descriptor Value Length        WORD    2
'Descriptor Value             See text  varies
Set tempNode = TV.Nodes.Add("ExtCont", tvwChild, "ExtContData", "Extended Content Data", 3)
ExtContSize = 0
For i = MaxTags To DescriptorCount - 1
    TagIndex = i ' MaxTags + i
    With Descriptor(TagIndex)
        Get #fNum, , .NameLen 'Name Length
        ExtContSize = ExtContSize + LenB(.NameLen)
        ExtContSize = ExtContSize + .NameLen
        ReDim .bytTagName(.NameLen - 1)
        'get the name of the tag
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        'Debug.Print LastAddress$; " Before TagName "; i
        Get #fNum, , .bytTagName
        'convert the byte data to a string by forcing it into the string value
        .Name = .bytTagName
        'trim the null from the enc
        .Name = Left$(.Name, Len(.Name) - 1)
        'get the Data Type of the tag
        Get #fNum, , .DataType  '0 indicates a string
        ExtContSize = ExtContSize + LenB(.DataType)
        'get the length of the Tag Data
        Get #fNum, , .DataLen
        ExtContSize = ExtContSize + LenB(.DataLen)
        ExtContSize = ExtContSize + .DataLen
        If .DataLen > 0 Then
            ReDim .bytTagData(.DataLen - 1)
            'get the tag data
            Get #fNum, , .bytTagData
        '0x0000  Unicode string The data consists of a sequence of Unicode characters.
        '0x0001  BYTE array The type of data is implementation-specific.
        'i think bool is 4 bytes long in extended content descr
        '0x0002  BOOL.      The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
        '                   Only 0x0000 or 0x0001 are permitted values.
        '0x0003  DWORD.     The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
        '0x0004  QWORD.     The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
        '0x0005  WORD.      The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
        '
            'the Tag data must be set according to its datatype
            Dim tmpLong As Long
            Select Case .DataType
                Case Is = 0 'string values
                    .Data = .bytTagData
                    .Data = Left$(.Data, Len(.Data) - 1)
                Case Is = 1  'BYTE array
                    .Data = .bytTagData
                Case Is = 2 'Supposed to be Boolean!!  VB boolean is 2 bytes - MetaData is 4 bytes
                    CopyMemory ByVal VarPtr(tmpLong), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                    .Data = CBool(tmpLong)
                Case Is = 3 'Long
                    CopyMemory ByVal VarPtr(tmpLong), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                    .Data = tmpLong
                Case Is = 4 'Date or Double...probably date, but i can't figure out how to convert it
                    'Date variables are stored as IEEE 64-bit (8-byte) floating-point numbers that
                    'represent dates ranging from 1 January 100 to 31 December 9999
                    'and times from 0:00:00 to 23:59:59.
                    'Any recognizable literal date values can be assigned to Date variables.
                    'When othernumeric types are converted to Date, values to the left of the decimal
                    'represent date information while values to the right of the decimal represent time.
                    'Midnight is 0 and midday is 0.5.
                    'Negative whole numbers represent dates before 30 December 1899.
                    Dim tmpDouble As Double
                    Dim tmpDate As Long
                    Dim tmpTime As Long
                    Dim tmpD As Date
                    CopyMemory ByVal VarPtr(tmpDouble), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                    CopyMemory ByVal VarPtr(tmpDate), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen / 2)
                    CopyMemory ByVal VarPtr(tmpTime), ByVal VarPtr(.bytTagData(.DataLen / 2)), CLng(.DataLen / 2)
                    .Data = tmpDouble
                    Dim HexStr$
                    For j = 0 To .DataLen - 1
                        HexStr$ = Right$("0" & (Hex$(.bytTagData(j))), 2) & HexStr$
                    Next j
                    tmpDouble = CDbl(Val("&H" & HexStr$))
                    tmpDate = CLng("&H" & Left(HexStr$, 8))
                    tmpTime = CLng("&H" & Right(HexStr$, 8))
                    'tmpD = CDate(tmpDouble) 'CLng("&H" & Left(hexstr$, 8)))
                Case Is = 5 'integer
                    Dim tmpInteger As Integer
                    CopyMemory ByVal VarPtr(tmpInteger), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                    .Data = tmpInteger
                Case Else
                    Dim wmaVariant As Variant
                    CopyMemory (wmaVariant), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                    .Data = wmaVariant
            End Select
        End If
        strInfo$ = LastAddress$ & ": " & .Name & " = " & .Data
        eM$ = eM$ & strInfo$ & vbCrLf
        Set tempNode = TV.Nodes.Add("ExtContData", tvwChild, "EC_Data" & TagIndex, TagIndex & ")" & strInfo$, 4)
        Debug.Print "'"; """"; .Name; " = "; .Data; """"
        Debug.Print ".NameLen = "; .NameLen; " '("; Hex$(.NameLen); ")"
        Debug.Print ".Name = "; """"; .Name; """"
        Debug.Print "'VB Data Type = "; VarType(.Data); " => "; TypeName(.Data)
        Debug.Print ".DataType = "; .DataType
        Debug.Print ".DataLen = "; .DataLen; " '("; Hex$(.DataLen); ")"
        Debug.Print ".Data = "; .Data
        Debug.Print "ExtContSize ="; ExtContSize; " '("; Hex$(ExtContSize); ")"
        Debug.Print "-------------End of "; i; "--------------------"
    End With
    With InfoTag
        Select Case Descriptor(TagIndex).Name
            'if we want, we can add some of these as properties, but in this
            'example, i just want the basics
            Case "WM/Track"
            Case "WM/Lyrics"
            Case "WM/MediaPrimaryClassID"
            Case "WMFSDKVersion"
            Case "WMFSDKNeeded"
            Case "IsVBR"
            Case "WM/TrackNumber"
                .TrackNumber = Descriptor(TagIndex).Data
            Case "WM/Year"
                .Year = Descriptor(TagIndex).Data
            Case "WM/EncodingTime"
            Case "WM/UniqueFileIdentifier"
            Case "WM/Composer"
                .Composer = Descriptor(TagIndex).Data
            Case "WM/Publisher"
                .Publisher = Descriptor(TagIndex).Data
            Case "WM/Genre"
                .Genre = Descriptor(TagIndex).Data
            Case "WM/AlbumTitle"
                .Album = Descriptor(TagIndex).Data
            Case "WM/AlbumArtist"
                .AlbumArtist = Descriptor(TagIndex).Data
            Case "WM/MCDI"
            Case "WM/Provider"
            Case "WM/ProviderRating"
                .ProviderRating = Descriptor(TagIndex).Data
            Case "WM/ProviderStyle"
        End Select
    End With
Next i
For i = MaxTags To DescriptorCount - 1
    With Descriptor(i)
        Debug.Print "Ext. Content for"; i; " = "; .Name; " => "; .DataLen; " = "; .Data
    End With
Next i
MaxTags = MaxTags + DescriptorCount
ExtContSize = ExtContSize + 26
Debug.Print
Debug.Print "ExtContSize   ="; ExtContSize; " ("; Hex$(ExtContSize); ")"
Debug.Print "ExtCont_QSize ="; ExtCont_QSize; " ("; Hex$(ExtCont_QSize); ")"
'
Debug.Print "----------------------------------------------------------"
Debug.Print eM$
Debug.Print "----------------------------------------------------------"
'MsgBox eM$, , "Tag Data - Tag Count = " & DescriptorCount
'
'this will give us the following data:
'Dear Johnny
'Poe
'1,042,973 (FEA1D)   73(49)  66(42)  206(CE)
'
'WM/Track = 10
'WM/Lyrics =
'WM/MediaPrimaryClassID = {D1607DBC-E323-4BE2-86A1-48A42A28441E}
'WMFSDKVersion = 10.00.00.3646
'WMFSDKNeeded = 0.0.0.0000
'IsVBR = 0
'WM/TrackNumber = 11
'WM/Year = 2000
'WM/EncodingTime = 3.93107752759597E-300
'WM/UniqueFileIdentifier = AMGa_id=R   498769;AMGp_id=P   169531;AMGt_id=T  4227544
'WM/Composer = Poe
'WM/Publisher = Atlantic
'WM/Genre = Rock
'WM/AlbumTitle = Haunted
'WM/AlbumArtist = Poe
'WM/MCDI = 12+B6+15A1+7373+DDE0+1304E+18560+194BA+23309+2717F+2B2A5+2F4BF+30381+36198+39D02+3F2DE+412A8+48305+4E0D1+52316
'WM/Provider = AMG
'WM/ProviderRating = 6
'WM/ProviderStyle = Rock
'
GoTo Exit_ParseExtContentData
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseExtContentData "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseExtContentData"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseExtContentData:
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ParseExtContentData"
End Sub

Sub ParseFileProperties()
On Error GoTo Oops
'------------------------------ASF_File_Properties_Object----------------------------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ParseFileProperties"
Get #fNum, , objFileInfo
HeaderCompleteCount = HeaderCompleteCount + 1
GUIDIn = objFileInfo.ObjectID
strFP_GUID = GUIDToString(GUIDIn)
Debug.Print "-------------------ASF_File_Properties_Object-----------------------------"
Debug.Print "File GUID = "; strFP_GUID
eM$ = "Executing " & ReturnGUIDName(strFP_GUID) & vbCrLf & strFP_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strFP_GUID <> ASF_File_Properties_Object Then
    eM$ = eM$ & vbCrLf & "** Error in strFP_GUID = " & vbCrLf & strFP_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_File_Properties_Object"
End If
'
Dim DateCreated As Date
Dim varDuration As Variant
Dim varSendDuration As Variant
With objFileInfo
    FP_QSize = QWORD(.ObjectSize) ' Valid values are at least 104 bytes
    Accum_HeaderQSize = Accum_HeaderQSize + FP_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseFileProperties"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    Debug.Print "ObjectSize             = "; FP_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    GUIDIn = .FileID
    strFile_GUID = GUIDToString(GUIDIn) 'identical to File ID in Data Object
    Debug.Print "File GUID = "; strFile_GUID
'
    FileQSize = QWORD(.FileSize) 'value of this field is invalid if the Broadcast Flag is set to 1.
    Debug.Print "FileSize               = "; FileQSize; "("; Hex$(.FileSize.High); AddZero(Hex$(.FileSize.Low), 8); ")"
'
    'Specifies the date and time of the initial creation of the file.
    'The value is given as the number of 100-nanosecond intervals since January 1, 1601,
    'according to Coordinated Universal Time (Greenwich Mean Time).
    CreationDateQSize = QWORD(.CreationDate)
    Debug.Print "CreationDate Intervals = "; Format(CreationDateQSize, "###,#"); "("; Hex$(.CreationDate.High); AddZero(Hex$(.CreationDate.Low), 8); ")"
    'Extract date
    DateCreated = ExtractDate(CreationDateQSize)
    Debug.Print "CreationDate           = "; DateCreated
    'get data packet size
    DataPacketsQSize = QWORD(.DataPackets)
    Debug.Print "DataPackets            = "; DataPacketsQSize; "("; Hex$(.DataPackets.High); AddZero(Hex$(.DataPackets.Low), 8); ")"
    'get the play duration
    PlayDurationQSize = QWORD(.PlayDuration)
    Debug.Print "PlayDuration           = "; PlayDurationQSize; "("; Hex$(.PlayDuration.High); AddZero(Hex$(.PlayDuration.Low), 8); ")"
    varDuration = CVar(PlayDurationQSize * ((10 ^ -7)))
    InfoTag.Length = varDuration
    InfoTag.Duration = fTime(InfoTag.Length)
    Debug.Print "Play Duration time     = "; varDuration; " sec.", InfoTag.Duration
    'get the send duration
    SendDurationQSize = QWORD(.SendDuration)
    Debug.Print "SendDuration           = "; SendDurationQSize; "("; Hex$(.SendDuration.High); AddZero(Hex$(.SendDuration.Low), 8); ")"
    varSendDuration = CVar(SendDurationQSize * ((10 ^ -7)))
    Debug.Print "Send Duration time     = "; ; varSendDuration; " sec."
    'get the PreRoll
    PreRollQSize = QWORD(.Preroll)
    Debug.Print "PreRoll                = "; PreRollQSize; "("; Hex$(.Preroll.High); AddZero(Hex$(.Preroll.Low), 8); ")"
'
    Debug.Print "Flags                  = "; .Flags; "("; Hex$(.Flags); ")"
    'Broadcast Flag (bit 0)
    'SeekableFlag   (bit 1)
    BroadcastFlag = (.Flags And 1) / 1
    SeekableFlag = (.Flags And 2) / 2
    Debug.Print "BroadcastFlag (bit 0)  = "; BroadcastFlag
    Debug.Print "SeekableFlag  (bit 1)  = "; SeekableFlag
    Debug.Print "Reserved (30 bits)     = "; (.Flags And 65532)
    Debug.Print "MinPacketSize          = "; .MinPacketSize; "("; Hex$(.MinPacketSize); ")"
    Debug.Print "MaxPacketSize          = "; .MaxPacketSize; "("; Hex$(.MaxPacketSize); ")"
    Debug.Print "MaxBitrate             = "; .MaxBitrate; "("; Hex$(.MaxBitrate); ")"
    '
    If BroadcastFlag = 1 Then
        MsgBox "BroadcastFlag = 1"
    End If
    If (FileQSize <> InfoTag.FileSize) And BroadcastFlag = 0 Then
        eM$ = eM$ & vbCrLf & "*** Error in FileSize ***"
        eM$ = eM$ & vbCrLf & "Actual FileSize = " & InfoTag.FileSize & " (" & Hex$(InfoTag.FileSize) & ")"
        eM$ = eM$ & vbCrLf & "Object FileSize = " & FileQSize & " (" & Hex$(FileQSize) & ")"
        eM$ = eM$ & vbCrLf & "Difference = " & InfoTag.FileSize - FileQSize & " (" & Hex$(InfoTag.FileSize - FileQSize) & ")"
        'frmMain.txtStatus.SelText = eM$ & vbCrLf
        eM$ = ""
    End If
    'now put our data into the tree
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "FP", LastAddress$ & ": " & "ASF_File_Properties_Object", 3)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "FP_GUID", "GUID = " & strFP_GUID, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "FP_Size", "ObjectSize     = " & FP_QSize & " (" & Hex$(FP_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "File_GUID", "FileGUID       = " & strFile_GUID, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "File_Size", "File Size      = " & FileQSize & " (" & Hex$(FileQSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "CreationDate", "CreationDate   = " & DateCreated, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "DataPackets", "DataPackets    = " & DataPacketsQSize, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "PlayDuration", "Play Duration  = " & varDuration, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "SendDuration", "Send Duration  = " & varSendDuration, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "Preroll", "Preroll        = " & PreRollQSize & " (" & Hex$(PreRollQSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "BroadcastFlag", "BroadcastFlag  = " & BroadcastFlag, 4)
    Set tempNode = TV.Nodes.Add("FP", tvwChild, "SeekableFlag", "SeekableFlag    = " & SeekableFlag, 4)
End With
If Len(eM$) > 0 Then
    Debug.Print eM$
    MsgBox eM$, , "ASF_File_Properties_Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ParseFileProperties"
GoTo Exit_ParseFileProperties
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseFileProperties "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseFileProperties"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseFileProperties:
End Sub

Sub ParseHeaderExtension()
On Error GoTo Oops
'--------------------------------ASF_Header_Extension_Object---------------------------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objHeaderExtInfo"
Get #fNum, , objHeaderExtInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objHeaderExtInfo.ObjectID
strEXT_GUID = GUIDToString(GUIDIn)
Debug.Print "-------------------ASF_Header_Extension_Object-----------------------------"
Debug.Print "ObjectID = "; strEXT_GUID
eM$ = "Executing " & ReturnGUIDName(strEXT_GUID) & vbCrLf & strEXT_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strEXT_GUID <> ASF_Header_Extension_Object Then
    eM$ = "ASF_Header_Extension_Object"
    eM$ = eM$ & vbCrLf & "---------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strEXT_GUID = " & vbCrLf & strEXT_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
'
With objHeaderExtInfo
    Ext_QSize = QWORD(.ObjectSize)
    'don't add this to the accum. value as it encompasses several other objects
    Accum_HeaderQSize = Accum_HeaderQSize + Len(objHeaderExtInfo) ' Ext_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseHeaderExtension"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    Debug.Print "ObjectSize         = "; Ext_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    GUIDIn = .Reserved1
    strRes1_GUID = GUIDToString(GUIDIn)
    If strRes1_GUID <> ASF_Reserved_1 Then
        eM$ = eM$ & vbCrLf & "** Error in ASF_Reserved_1 = " & strRes1_GUID
        Debug.Print eM$
        MsgBox eM$
        eM$ = ""
    Else
        Debug.Print "Reserved1 GUID OK  = "; strRes1_GUID
    End If
    If .Reserved2 <> 6 Then 'always 6
        eM$ = eM$ & vbCrLf & "** Error in Reserved2 = " & .Reserved2 & " (" & Hex$(.Reserved2) & ")"
    Else
        Debug.Print "Reserved2 OK       = "; .Reserved2; " - always 6"
    End If
    '
    'Specifies the number of bytes stored in the Header Extension Data field.
    'This value may be 0 bytes or 24 bytes and larger.
    'It should also be equal to the Object Size field minus 46 bytes.
    ReDim bytExtData(.ExtDataSize - 1)
    ExtDataBytesLeft = .ExtDataSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft & " (" & Hex$(.ExtDataSize) & ")"
    If .ExtDataSize <> Ext_QSize - 46 Then
        eM$ = "There may be an error with the size of the extended data..."
        eM$ = eM$ & vbCrLf & "ObjectSize    = " & Ext_QSize & " (" & Hex$(Ext_QSize) & ")"
        eM$ = eM$ & vbCrLf & "ExtDataSize   = " & .ExtDataSize & " (" & Hex$(.ExtDataSize) & ")"
    Else
        Debug.Print "ExtDataSize OK     = "; .ExtDataSize; "("; Hex$(.ExtDataSize); ") = Ext_QSize - 46 "
    End If
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "HE", LastAddress$ & ": " & "ASF_Header_Extension_Object", 3)
    Set tempNode = TV.Nodes.Add("HE", tvwChild, "strEXT_GUID", "GUID = " & strEXT_GUID, 4)
    Set tempNode = TV.Nodes.Add("HE", tvwChild, "Ext_QSize", "ObjectSize    = " & Ext_QSize & " (" & Hex$(Ext_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("HE", tvwChild, "strRes1_GUID", "strRes1_GUID  = " & strRes1_GUID, 4)
    Set tempNode = TV.Nodes.Add("HE", tvwChild, "HE_Reserved2", "Reserved2     = " & .Reserved2, 4)
    Set tempNode = TV.Nodes.Add("HE", tvwChild, "ExtDataSize", "ExtDataSize   = " & .ExtDataSize & " (" & Hex$(.ExtDataSize) & ")", 4)
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Header_Extension_Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objHeaderExtInfo"
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " at beginning of ExtDataBytesLeft"
GoTo Exit_ParseHeaderExtension
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseHeaderExtension "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseHeaderExtension"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseHeaderExtension:
End Sub

'=========================================================================
'Below data is all part of the extended header
'they don't count as header objects
'ParseCompatibility
'ParseMetaData
'ParsePadding
'ParseExtendedStream
'ParsePlaceholder
'ParseMetadataLibrary
'=========================================================================

Sub ParseCompatibility()
On Error GoTo Oops
'-------------------------------Undocumented Compatibility Object---------------------------
'Undocumented object...
'I could not find the guid anywhere, so i reverse engineered the structure to verify thatj
'the reserved compatibility object is the same.
'
'Private Type CompatibilityObject '(26 bytes)
'    ObjectID As GuidStruct  '16 bytes
'    ObjectSize As DlngSize  '8 bytes The value of this field shall be set to 26 bytes
'    Profile As Byte         '1 byte 'This field is reserved and should be set to 2.
'    Mode As Byte            '1 byte 'This field is reserved and should be set to 1.
'End Type
'let's try to get the Compatibility data
'first get the guid
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objCompatibilityInfo"
Get #fNum, , objCompatibilityInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objCompatibilityInfo.ObjectID
strComp_GUID = GUIDToString(GUIDIn)
Debug.Print "GUID = "; strComp_GUID
eM$ = "Executing " & ReturnGUIDName(strComp_GUID) & vbCrLf & strComp_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
'
If strComp_GUID <> ASF_Undocumented_Compatibility_GUID Then
    eM$ = "Compatibility Object"
    eM$ = eM$ & vbCrLf & "---------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strComp_GUID = " & vbCrLf & strComp_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
'-------------------------------------------------------------------------------------
With objCompatibilityInfo
    Comp_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + Comp_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseCompatibility"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    ExtDataBytesLeft = ExtDataBytesLeft - Comp_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    'The value of this field shall be set to 26 bytes
    If Comp_QSize <> 26 Then
        eM$ = eM$ & vbCrLf & "** Error in ObjectSize = " & Comp_QSize & " (" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ") - always 26"
    Else
        Debug.Print "ObjectSize OK  = "; Comp_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ") - always 26"
    End If
    If .Profile <> 2 Then 'always 2
        eM$ = eM$ & vbCrLf & "** Error in Profile = " & .Profile & " (" & Hex$(.Profile) & ")"
    Else
        Debug.Print "Profile OK     = "; .Profile; "("; Hex$(.Profile); ") - always 2"
    End If
    If .Mode <> 1 Then  'always 1
        eM$ = eM$ & vbCrLf & "** Error in Mode = " & .Mode & " (" & Hex$(.Mode) & ")"
    Else
        Debug.Print "Mode OK        = "; .Mode; "("; Hex$(.Mode); ") - always 1"
    End If
    '
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Compat", LastAddress$ & ": " & "Compatibility Object", 3)
    Set tempNode = TV.Nodes.Add("Compat", tvwChild, "strComp_GUID", "GUID = " & strComp_GUID, 4)
    Set tempNode = TV.Nodes.Add("Compat", tvwChild, "Comp_QSize", "ObjectSize = " & Comp_QSize & " (" & Hex$(Comp_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Compat", tvwChild, "Profile", "Profile    = " & .Profile, 4)
    Set tempNode = TV.Nodes.Add("Compat", tvwChild, "Mode", "Mode       = " & .Mode, 4)
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "Compatibility Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objCompatibilityInfo"
GoTo Exit_ParseCompatibility
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseCompatibility "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseCompatibility"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseCompatibility:
End Sub

Sub ParseMetaData()
On Error GoTo Oops
'--------------------------------ASF_Metadata_Object---------------------------------------
'Private Type MetaDataObject '(26 bytes)
'    ObjectID As GuidStruct  '16 bytes =>ASF_Metadata_Object
'    ObjectSize As DlngSize  '8 bytes =>The value of this field shall be set to 26 bytes
'    RecordCount As Integer 'Description Records Count
'End Type
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objMetaDataInfo"
Get #fNum, , objMetaDataInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objMetaDataInfo.ObjectID
strMetaData_GUID = GUIDToString(GUIDIn)
Debug.Print "-------------------ASF_Metadata_Object-----------------------------"
Debug.Print "MetaData GUID = "; strMetaData_GUID
eM$ = "Executing " & ReturnGUIDName(strMetaData_GUID) & vbCrLf & strMetaData_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strMetaData_GUID <> ASF_Metadata_Object Then
    eM$ = eM$ & vbCrLf & "** Error in strMetaData_GUID = " & vbCrLf & strMetaData_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_Metadata_Object"
End If
With objMetaDataInfo
    MetaData_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + MetaData_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseMetaData"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    ExtDataBytesLeft = ExtDataBytesLeft - MetaData_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize = "; MetaData_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    Debug.Print "Meta Data  = "; MetaData_QSize - 26
    Debug.Print "RecordCount= "; .RecordCount; "("; Hex$(.RecordCount); ")"
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Meta", LastAddress$ & ": " & "ASF_Metadata_Object", 3)
    Set tempNode = TV.Nodes.Add("Meta", tvwChild, "strMetaData_GUID", "GUID = " & strMetaData_GUID, 4)
    Set tempNode = TV.Nodes.Add("Meta", tvwChild, "MetaData_QSize", "ObjectSize    = " & MetaData_QSize & " (" & Hex$(MetaData_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Meta", tvwChild, "MetaRecordCount", "RecordCount   = " & .RecordCount, 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objMetaDataInfo"
'
Set tempNode = TV.Nodes.Add("Meta", tvwChild, "MetaData", LastAddress$ & ": " & "MetaData Entries", 3)
ParseMetaDataData "MetaData", objMetaDataInfo.RecordCount
GoTo Exit_ParseMetaData
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseMetaData "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseMetaData"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseMetaData:
End Sub

Sub ParseMetaDataData(RootName As String, MetaDataCount As Integer)
On Error GoTo Oops
Dim strInfo$
ReDim Preserve MetaDescriptor(MaxMTags + MetaDataCount)
eM$ = ""
Debug.Print "Retrieving data for "; MetaDataCount; " MetaData Entries."
'
Debug.Print "----------------------------------------------------------"
Dim tmpLong As Long
For i = MaxMTags To MaxMTags + MetaDataCount - 1
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print LastAddress$; " Before ParseMetaDataData"
    Get #fNum, , objMetaDataRecord
    MTagIndex = i
    With MetaDescriptor(MTagIndex)
        .Reserved = objMetaDataRecord.Reserved
        .StreamNumber = objMetaDataRecord.StreamNumber
        'get Length of the name data
        .NameLen = objMetaDataRecord.NameLen
        .DataLen = objMetaDataRecord.DataLen
        .DataType = objMetaDataRecord.DataType
        If .Reserved <> 0 Then 'always 0
            eM$ = eM$ & vbCrLf & "** Error in Reserved = " & .Reserved & " (" & Hex$(.Reserved) & ")"
        Else
            Debug.Print "Reserved OK     = "; .Reserved; "("; Hex$(.Reserved); ") - always 0"
        End If
        Debug.Print "StreamNumber    = "; .StreamNumber; "("; Hex$(.StreamNumber); ")"
        'A value of 0 in this field indicates that it applies to the whole file;
        ' otherwise, the entry applies only to the indicated stream number
        ' and must be between 1 and 127.
        '
        Debug.Print "MetaData NameLen = "; .NameLen
        ReDim .bytTagName(.NameLen - 1)
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before bytTagName"
        'get the name of the MTag
        Get #fNum, , .bytTagName
        'convert the byte data to a string by forcing it into the string value
        .Name = .bytTagName
        'trim the null from the Name
        .Name = Left$(.Name, Len(.Name) - 1)
        Debug.Print .Name
        '
        'get the data type
        Debug.Print "MetaData DataType = "; .DataType; " ("; Hex$(.DataType); ")"
        'get the data length
        Debug.Print "MetaData DataLen = "; .DataLen; " ("; Hex$(.DataLen); ")"
        '
        'Get the MetaData Data
        If .DataLen > 0 Then
            If .DataType = 6 Then 'guid
                GetGUID
                .Data = strGUID
                GoTo ShowMe
            Else
                ReDim .bytTagData(.DataLen - 1)
                'get the Data of the MTag
                Get #fNum, , .bytTagData
                'convert the byte data to a string by forcing it into the string value
                .Data = .bytTagData
                'trim the null from the Data
                '.data = Left$(.data, Len(.data) - 1)
            End If
        End If
        'Metadata Library data type
'        0x0000  Unicode string. The data consists of a sequence of Unicode characters.
'        0x0001  BYTE array. The type of the data is implementation-specific.
'        0x0002  BOOL. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'                      Only 0x0000 or 0x0001 are permitted values.
'        0x0003  DWORD. The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'        0x0004  QWORD. The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'        0x0005  WORD. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'        0x0006  GUID. The data is 16 bytes long and should be interpreted as a 128-bit GUID.
       'the Tag data must be set according to its datatype
        Select Case .DataType
            Case Is = 0 'string values
                .Data = .bytTagData
                .Data = Left$(.Data, Len(.Data) - 1)
            Case Is = 1  'BYTE array
                .Data = .bytTagData
                .Data = Left$(.Data, Len(.Data) - 1)
            Case 2 'Boolean
                CopyMemory ByVal VarPtr(tmpLong), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                .Data = CBool(tmpLong)
            Case 3 'Long
                CopyMemory ByVal VarPtr(tmpLong), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                .Data = tmpLong
            Case 4 'Date or Double...probably date, but i can't figure out how to convert it
                'Date variables are stored as IEEE 64-bit (8-byte) floating-point numbers that
                'represent dates ranging from 1 January 100 to 31 December 9999
                'and times from 0:00:00 to 23:59:59.
                'Any recognizable literal date values can be assigned to Date variables.
                'When othernumeric types are converted to Date, values to the left of the decimal
                'represent date information while values to the right of the decimal represent time.
                'Midnight is 0 and midday is 0.5.
                'Negative whole numbers represent dates before 30 December 1899.
                Dim tmpDouble As Double
                Dim tmpDate As Long
                Dim tmpTime As Long
                Dim tmpD As Date
                CopyMemory ByVal VarPtr(tmpDouble), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                CopyMemory ByVal VarPtr(tmpDate), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen / 2)
                CopyMemory ByVal VarPtr(tmpTime), ByVal VarPtr(.bytTagData(.DataLen / 2)), CLng(.DataLen / 2)
                .Data = tmpDouble
                Dim HexStr$
                For j = 0 To .DataLen - 1
                    HexStr$ = Right$("0" & (Hex$(.bytTagData(j))), 2) & HexStr$
                Next j
                tmpDouble = CDbl(Val("&H" & HexStr$))
                tmpDate = CLng("&H" & Left(HexStr$, 8))
                tmpTime = CLng("&H" & Right(HexStr$, 8))
                'tmpD = CDate(tmpDouble) 'CLng("&H" & Left(hexstr$, 8)))
            Case 5 'integer
                Dim tmpInteger As Integer
                CopyMemory ByVal VarPtr(tmpInteger), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                .Data = tmpInteger
            Case 6 ' GUID. The data is 16 bytes long and should be interpreted as a 128-bit GUID.
                MsgBox "Get a Guid in ParseMetaDataData"
             Case Else
                Dim wmaVariant As Variant
                CopyMemory (wmaVariant), ByVal VarPtr(.bytTagData(0)), CLng(.DataLen)
                .Data = wmaVariant
        End Select
ShowMe:
        'Show the tag name and data
        strInfo$ = .Name & " = " & .Data
        Debug.Print strInfo$
        'eM$ = eM$ & strInfo$ & vbCrLf
        Set tempNode = TV.Nodes.Add(RootName, tvwChild, RootName & MTagIndex, LastAddress$ & ": " & MTagIndex & "=>" & .Name, 3)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MReserved" & MTagIndex, "Reserved     =" & objMetaDataRecord.Reserved, 4)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MStreamNumber" & MTagIndex, "StreamNumber =" & objMetaDataRecord.StreamNumber, 4)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MNameLen" & MTagIndex, "NameLen      =" & .NameLen, 4)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MName" & MTagIndex, "Name         =" & .Name, 4)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MDataLen" & MTagIndex, "DataLen      =" & .DataLen, 4)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MDataType" & MTagIndex, "DataType     =" & .DataType, 4)
        Set tempNode = TV.Nodes.Add(RootName & MTagIndex, tvwChild, "MData" & MTagIndex, "Data         =" & .Data, 4)
        'set the data if we need to...
        Select Case .Name
            Case "WM/TrackNumber"
                InfoTag.TrackNumber = .Data
        End Select
    End With
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print LastAddress$; " After objMetaDataRecord " & i
Next i
Debug.Print "--------------MetaData Data------------------------"
For i = MaxMTags To MaxMTags + MetaDataCount - 1
    With MetaDescriptor(i)
        Debug.Print "MetaData for"; i; " = "; .Name; " = "; .Data
    End With
Next i
MaxMTags = MaxMTags + MetaDataCount
Debug.Print "----------------------------------------------------------"
'
GoTo Exit_ParseMetaDataData
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseMetaDataData "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseMetaDataData"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseMetaDataData:
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ParseMetaDataData"
End Sub

Sub ParsePadding()
On Error GoTo Oops
'--------------------------------ASF_Padding_Object---------------------------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objPaddingInfo"
Get #fNum, , objPaddingInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objPaddingInfo.ObjectID
strPad_GUID = GUIDToString(GUIDIn)
'should be 1806D474-CADF-4509-A4BA-9AABCB96AAE8
Debug.Print "-------------------ASF_Padding_Object-----------------------------"
Debug.Print "objPaddingInfo GUID = "; strPad_GUID
eM$ = "Executing " & ReturnGUIDName(strPad_GUID) & vbCrLf & strPad_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strPad_GUID <> ASF_Padding_Object Then
    eM$ = eM$ & vbCrLf & "** Error in strPad_GUID = " & vbCrLf & strPad_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_Padding_Object"
End If
With objPaddingInfo
    Pad_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + Pad_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParsePadding"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    ExtDataBytesLeft = ExtDataBytesLeft - Pad_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize = "; Pad_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    PaddingLen = Pad_QSize - 24
    Debug.Print "PaddingLen = "; PaddingLen; "("; Hex$(PaddingLen); ")"
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Padding", LastAddress$ & ": " & "ASF_Padding_Object", 3)
    Set tempNode = TV.Nodes.Add("Padding", tvwChild, "strPad_GUID", "GUID = " & strPad_GUID, 4)
    Set tempNode = TV.Nodes.Add("Padding", tvwChild, "Pad_QSize", "ObjectSize  = " & Pad_QSize & " (" & Hex$(Pad_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Padding", tvwChild, "PaddingLen", "PaddingLen  = " & PaddingLen & " (" & Hex$(PaddingLen) & ")", 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objPaddingInfo"
ReDim bytPaddingData(PaddingLen - 1)
Get #fNum, , bytPaddingData
PaddingEndPos = Seek(fNum)
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After bytPaddingData"
GoTo Exit_ParsePadding
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParsePadding "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParsePadding"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParsePadding:
End Sub

Sub ParseExtendedStream()
On Error GoTo Oops
'------------------------ASF_Extended_Stream_Properties_Object-----------------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objExtStreamInfo"
Get #fNum, , objExtStreamInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objExtStreamInfo.ObjectID
strExtStream_GUID = GUIDToString(GUIDIn)
'should be 14E6A5CB-C672-4332-8399-A96952065B5A
Debug.Print "-------------------ASF_Extended_Stream_Properties_Object-----------------------------"
Debug.Print "objExtStreamInfo GUID = "; strExtStream_GUID
'the next guid should be ASF_Extended_Stream_Properties_Object
eM$ = "Executing " & ReturnGUIDName(strExtStream_GUID) & vbCrLf & strExtStream_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strExtStream_GUID <> ASF_Extended_Stream_Properties_Object Then
    eM$ = eM$ & vbCrLf & "** Error in strExtStream_GUID = " & vbCrLf & strExtStream_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_Extended_Stream_Properties_Object"
End If
'
Dim varAvgTime As Variant
With objExtStreamInfo
    ExtStream_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + ExtStream_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseExtendedStream"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    ExtDataBytesLeft = ExtDataBytesLeft - ExtStream_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize             = "; ExtStream_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    StartTimeQSize = QWORD(.StartTime)
    Debug.Print "StartTime              = "; StartTimeQSize; "("; Hex$(.StartTime.High); AddZero(Hex$(.StartTime.Low), 8); ")"
    '
    EndTimeQSize = QWORD(.EndTime)
    Debug.Print "EndTime                = "; EndTimeQSize; "("; Hex$(.EndTime.High); AddZero(Hex$(.EndTime.Low), 8); ")"
    '
    Debug.Print "DataBitrate            = "; .DataBitrate; "("; Hex$(.DataBitrate); ")"
    Debug.Print "BufferSize             = "; .BufferSize; "("; Hex$(.BufferSize); ")"
    Debug.Print "InitialBufferFullness  = "; .InitialBufferFullness; "("; Hex$(.InitialBufferFullness); ")"
    Debug.Print "AltDataBitrate         = "; .AltDataBitrate; "("; Hex$(.AltDataBitrate); ")"
    Debug.Print "AltBufferSize          = "; .AltBufferSize; "("; Hex$(.AltBufferSize); ")"
    Debug.Print "AltInitBufferFullness  = "; .AltInitBufferFullness; "("; Hex$(.AltInitBufferFullness); ")"
    Debug.Print "MaxObjectSize          = "; .MaxObjectSize; "("; Hex$(.MaxObjectSize); ")"
    Debug.Print "Flags                  = "; .Flags; "("; Hex$(.Flags); ")"
    '   Reliable Flag (LSB)
    '   Seekable Flag
    '   No Cleanpoints Flag
    '   Resend Live Cleanpoints Flag
    '   Reserved Flags                          28 bits
    Debug.Print "Reliable (bit 0)       = "; (.Flags And 1) / 1
    Debug.Print "SeekableFlag  (bit 1)  = "; (.Flags And 2) / 2
    Debug.Print "No Cleanpoints (bit 2) = "; (.Flags And 4) / 4
    Debug.Print "ResendLive  (bit 3)    = "; (.Flags And 8) / 8
    Debug.Print "Reserved (28 bits)     = "; (.Flags And 65520)
    '
    Debug.Print "StreamNumber           = "; .StreamNumber; "("; Hex$(.StreamNumber); ")"
    Debug.Print "StreamLangIDIndex      = "; .StreamLangIDIndex; "("; Hex$(.StreamLangIDIndex); ")"
    '
    AvgTimePerFrameQSize = QWORD(.AvgTimePerFrame)
    Debug.Print "AvgTimePerFrame        = "; AvgTimePerFrameQSize; "("; Hex$(.AvgTimePerFrame.High); AddZero(Hex$(.AvgTimePerFrame.Low), 8); ")"
    varAvgTime = CVar(AvgTimePerFrameQSize * ((10 ^ -7)))
    Debug.Print "AvgTimePerFrame (sec.) = "; varAvgTime
    '
    Debug.Print "StreamNameCount        = "; .StreamNameCount; "("; Hex$(.StreamNameCount); ")"
    ReDim StreamName(.StreamNameCount)
    Debug.Print "PayloadExtSystemCount  = "; .PayloadExtSystemCount; "("; Hex$(.PayloadExtSystemCount); ")"
    ReDim ExtSystemInfo(.PayloadExtSystemCount)
    '
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "ExtStream", LastAddress$ & ": " & "ASF_Extended_Stream_Properties_Object", 3)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "strExtStream_GUID", "GUID = " & strExtStream_GUID, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "ExtStream_QSize", "ObjectSize             = " & ExtStream_QSize & " (" & Hex$(ExtStream_QSize) & ")", 4)
    '
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "StartTime", "StartTime              = " & StartTimeQSize, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "EndTime", "EndTime                = " & EndTimeQSize, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "DataBitrate", "DataBitrate            = " & .DataBitrate, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "BufferSize", "BufferSize             = " & .BufferSize & " (" & Hex$(.BufferSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "InitialBufferFullness", "InitialBufferFullness  = " & .InitialBufferFullness & " (" & Hex$(.InitialBufferFullness) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "AltDataBitrate", "AltDataBitrate         = " & .AltDataBitrate & " (" & Hex$(.AltDataBitrate) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "AltBufferSize", "AltBufferSize          = " & .AltBufferSize & " (" & Hex$(.AltBufferSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "AltInitBufferFullness", "AltInitBufferFullness  = " & .AltInitBufferFullness & " (" & Hex$(.AltInitBufferFullness) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "MaxObjectSize", "MaxObjectSize          = " & .MaxObjectSize & " (" & Hex$(.MaxObjectSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "Flags", "Flags                  = " & .Flags, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "Reliable (bit 0)", "Reliable (bit 0)       = " & (.Flags And 1) / 1, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "SeekableFlag  (bit 1)", "SeekableFlag  (bit 1)  = " & (.Flags And 2) / 2, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "No Cleanpoints (bit 2)", "No Cleanpoints (bit 2) = " & (.Flags And 4) / 4, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "ResendLive  (bit 3)", "ResendLive  (bit 3)    = " & (.Flags And 8) / 8, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "Reserved (28 bits)", "Reserved (28 bits)     = " & (.Flags And 65520), 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "StreamNumber", "StreamNumber           = " & .StreamNumber, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "StreamLangIDIndex", "StreamLangIDIndex      = " & .StreamLangIDIndex, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "AvgTimePerFrame", "AvgTimePerFrame        = " & AvgTimePerFrameQSize & " (" & Hex$(AvgTimePerFrameQSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "AvgTimePerFrame (sec.)", "AvgTimePerFrame (sec.) = " & varAvgTime, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "StreamNameCount", "StreamNameCount        = " & .StreamNameCount, 4)
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "PayloadExtSystemCount", "PayloadExtSystemCount  = " & .PayloadExtSystemCount, 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objExtStreamInfo"
'Stream Names
'The structure of each Stream Name entry is shown in the following table.
'Field name             Field type  Size (bytes)
'-----------------------------------------------------------------------------
'Language ID Index      WORD        2
'Stream Name Length     WORD        2
'Stream Name            WCHAR   varies
If objExtStreamInfo.StreamNameCount > 0 Then
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "StreamNames", "StreamNames", 4)
    For i = 0 To objExtStreamInfo.StreamNameCount - 1
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before ExtStreamName"
        Get #fNum, , LangIDIndex
        Get #fNum, , StreamNLen
        ReDim bytStreamName(StreamNLen)
        Get #fNum, , bytStreamName
        'below array stores the name info
        StreamName(i) = bytStreamName
        Debug.Print i; ") StreamName = "; StreamName(i)
        Set tempNode = TV.Nodes.Add("StreamNames", tvwChild, "StreamName" & i, LastAddress$ & ": " & StreamName(i), 4)
    Next i
End If
'
'Payload extension systems are described in the following table.
'Field name  Field type  Size (bits)
''-----------------------------------------------------------------------------
'Extension System ID             GUID    16
'Extension Data Size             WORD    2
'Extension System Info Length    DWORD   4
'Extension System Info           BYTE    varies
'below is an array for each ExtSystemInfo
If objExtStreamInfo.PayloadExtSystemCount > 0 Then
    Set tempNode = TV.Nodes.Add("ExtStream", tvwChild, "SystemInfo", "Extension System Info", 4)
    For i = 0 To objExtStreamInfo.PayloadExtSystemCount - 1
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before objPayloadInfo"
        Get #fNum, , objPayloadInfo
        With objPayloadInfo
            GUIDIn = .ExtSystemID
            strExtSystem_GUID = GUIDToString(GUIDIn)
            'should be one of the following
'ASF_Payload_Extension_System_Timecode'             399595EC-8667-4E2D-8FDB-98814CE76C1E
'ASF_Payload_Extension_System_File_Name             E165EC0E-19ED-45D7-B4A7-25CBD1E28E9B
'ASF_Payload_Extension_System_Content_Type          D590DC20-07BC-436C-9CF7-F3BBFBF1A4DC
'ASF_Payload_Extension_System_Pixel_Aspect_Ratio    1B1EE554-F9EA-4BC8-821A-376B74E4C4B8
'ASF_Payload_Extension_System_Sample_Duration       C6BD9450-867F-4907-83A3-C77921B733AD
'ASF_Payload_Extension_System_Encryption_Sample_ID  6698B84E-0AFA-4330-AEB2-1C0A98D7A44D
            Debug.Print "objPayloadInfo GUID = "; strExtSystem_GUID
            Debug.Print "ExtDataSize        = "; .ExtDataSize; "("; Hex$(.ExtDataSize); ")"
            Debug.Print "ExtSystemInfoLen   = "; .ExtSystemInfoLen; "("; Hex$(.ExtSystemInfoLen); ")"
            'below gets the bytes for each ExtSystemInfo
            ReDim bytExtSystemInfo(.ExtSystemInfoLen - 1)
            Get #fNum, , bytExtSystemInfo
            ExtSystemInfo(i) = bytExtSystemInfo
            Debug.Print "ExtSystemInfo      = "; ExtSystemInfo(i); "("; Hex$(ExtSystemInfo(i)); ")"
            Set tempNode = TV.Nodes.Add("SystemInfo", tvwChild, "SystemInfo" & i, LastAddress$ & ": " & "SystemInfo" & i, 3)
            Set tempNode = TV.Nodes.Add("SystemInfo" & i, tvwChild, "strExtSystem_GUID", "GUID = " & strExtSystem_GUID, 4)
            Set tempNode = TV.Nodes.Add("SystemInfo" & i, tvwChild, "ExtDataSize", "ExtDataSize      =" & .ExtDataSize & " (" & Hex$(.ExtDataSize) & ")", 4)
            Set tempNode = TV.Nodes.Add("SystemInfo" & i, tvwChild, "ExtSystemInfoLen", "ExtSystemInfoLen =" & .ExtSystemInfoLen & " (" & Hex$(.ExtSystemInfoLen) & ")", 4)
            Set tempNode = TV.Nodes.Add("SystemInfo" & i, tvwChild, "ExtSystemInfo", "ExtSystemInfo    =" & ExtSystemInfo(i), 4)
        End With
    Next i
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objPayloadInfo"
Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
GoTo Exit_ParseExtendedStream
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseExtendedStream "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseExtendedStream"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseExtendedStream:
End Sub

Sub ParsePlaceholder()
On Error GoTo Oops
'---------------------------ASF_Index_Placeholder_Object----------------------------------------------------
'get the undocumented ASF_Index_Placeholder_Object
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ASF_Index_Placeholder_Object"
Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
Get #fNum, , objPlaceholderInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objPlaceholderInfo.ObjectID
strPH_GUID = GUIDToString(GUIDIn)
'should be ASF_Index_Placeholder_Object = "D9AADE20-7C17-4F9C-BC28-8555DD98E2A2"
Debug.Print "-------------------ASF_Index_Placeholder_Object-----------------------------"
Debug.Print "objPlaceholderInfo GUID = "; strPH_GUID
eM$ = "Executing " & ReturnGUIDName(strPH_GUID) & vbCrLf & strPH_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strPH_GUID <> ASF_Index_Placeholder_Object Then
    eM$ = "ASF_Index_Placeholder_Object"
    eM$ = eM$ & vbCrLf & "-----------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strPH_GUID = " & vbCrLf & strPH_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_Index_Placeholder_Object"
End If
With objPlaceholderInfo
    PH_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + PH_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParsePlaceholder"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    ExtDataBytesLeft = ExtDataBytesLeft - PH_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize = "; PH_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    If .Word1 <> 0 Then 'always 0
        eM$ = eM$ & vbCrLf & "** Error in Word1 = " & .Word1 & " (" & Hex$(.Word1) & ")"
    Else
        Debug.Print "Word1 OK = "; .Word1; "("; Hex$(.Word1); ")"
    End If
    If .Word2 <> 0 Then 'always 0
        eM$ = eM$ & vbCrLf & "** Error in Word2 = " & .Word2 & " (" & Hex$(.Word2) & ")"
    Else
        Debug.Print "Word2 OK = "; .Word2; "("; Hex$(.Word2); ")"
    End If
    If .Word3 <> 0 Then 'always 0
        eM$ = eM$ & vbCrLf & "** Error in Word3 = " & .Word3 & " (" & Hex$(.Word3) & ")"
    Else
        Debug.Print "Word3 OK = "; .Word3; "("; Hex$(.Word3); ")"
    End If
    If .Word4 <> 0 Then 'always 0
        eM$ = eM$ & vbCrLf & "** Error in Word4 = " & .Word4 & " (" & Hex$(.Word4) & ")"
    Else
        Debug.Print "Word4 OK = "; .Word4; "("; Hex$(.Word4); ")"
    End If
    If .Word5 <> 0 Then 'always 0
        eM$ = eM$ & vbCrLf & "** Error in Word5 = " & .Word5 & " (" & Hex$(.Word5) & ")"
    Else
        Debug.Print "Word5 OK = "; .Word5; "("; Hex$(.Word5); ")"
    End If
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Placeholder", LastAddress$ & ": " & "ASF_Index_Placeholder_Object", 3)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "strPH_GUID", "GUID = " & strPH_GUID, 4)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "PH_QSize", "ObjectSize = " & PH_QSize & " (" & Hex$(PH_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "PH_Word1", "Word1      = " & .Word1, 4)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "PH_Word2", "Word2      = " & .Word2, 4)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "PH_Word3", "Word3      = " & .Word3, 4)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "PH_Word4", "Word4      = " & .Word4, 4)
    Set tempNode = TV.Nodes.Add("Placeholder", tvwChild, "PH_Word5", "Word5      = " & .Word5, 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ASF_Index_Placeholder_Object"
Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
GoTo Exit_ParsePlaceholder
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParsePlaceholder "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParsePlaceholder"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParsePlaceholder:
End Sub

Sub ParseMetadataLibrary()
On Error GoTo Oops
'-----------------------------ASF_Metadata_Library_Object----------------------------------
'we also use this structure for ASF_Metadata_Object
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ASF_Metadata_Library_Object"
Get #fNum, , objMetaDataLibraryInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
GUIDIn = objMetaDataLibraryInfo.ObjectID
strMetaLibraryData_GUID = GUIDToString(GUIDIn)
Debug.Print "-------------------ASF_Metadata_Library_Object-----------------------------"
Debug.Print "Metadata_Library GUID = "; strMetaLibraryData_GUID
eM$ = "Executing " & ReturnGUIDName(strMetaLibraryData_GUID) & vbCrLf & strMetaLibraryData_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strMetaLibraryData_GUID <> ASF_Metadata_Library_Object Then
    eM$ = eM$ & vbCrLf & "** Error in strMetadata_Library_GUID = " & vbCrLf & strMetaLibraryData_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_Metadata_Library_Object"
End If
With objMetaDataLibraryInfo
    MetaDataLib_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + MetaDataLib_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseMetadataLibrary"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    ExtDataBytesLeft = ExtDataBytesLeft - MetaDataLib_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize = "; MetaDataLib_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    Debug.Print "RecordCount= "; .RecordCount; "("; Hex$(.RecordCount); ")"
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "MetaLibrary", LastAddress$ & ": " & "ASF_Metadata_Library_Object", 3)
    Set tempNode = TV.Nodes.Add("MetaLibrary", tvwChild, "strMetaDataLibrary_GUID", "GUID = " & strMetaLibraryData_GUID, 4)
    Set tempNode = TV.Nodes.Add("MetaLibrary", tvwChild, "MetaLibrary_QSize", "ObjectSize    = " & MetaDataLib_QSize & " (" & Hex$(MetaDataLib_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("MetaLibrary", tvwChild, "MetaLibraryRecCount", "RecordCount   = " & .RecordCount, 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ASF_Metadata_Library_Object"
'
Set tempNode = TV.Nodes.Add("MetaLibrary", tvwChild, "MetaDataLibrary", LastAddress$ & ": " & "MetaData Entries", 3)
ParseMetaDataData "MetaDataLibrary", objMetaDataLibraryInfo.RecordCount
GoTo Exit_ParseMetadataLibrary
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseMetadataLibrary "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseMetadataLibrary"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseMetadataLibrary:
End Sub

Sub ParseCodecList()
On Error GoTo Oops
'--------------------------------ASF_Codec_List_Object-------------------------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objCodecListInfo"
Get #fNum, , objCodecListInfo
HeaderCompleteCount = HeaderCompleteCount + 1
GUIDIn = objCodecListInfo.ObjectID
strCodec_GUID = GUIDToString(GUIDIn)
'should be 86D15240-311D-11D0-A3A4-00A0C90348F6
Debug.Print "-------------------ASF_Codec_List_Object-----------------------------"
Debug.Print "ObjectID = "; strCodec_GUID
eM$ = "Executing " & ReturnGUIDName(strCodec_GUID) & vbCrLf & strCodec_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
If strCodec_GUID <> ASF_Codec_List_Object Then
    eM$ = "ASF_Codec_List_Object"
    eM$ = eM$ & vbCrLf & "-----------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strCodec_GUID = " & vbCrLf & strCodec_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
With objCodecListInfo
    Codec_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + Codec_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseCodecList"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    Debug.Print "ObjectSize         = "; Codec_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    GUIDIn = .Reserved
    strRes2_GUID = GUIDToString(GUIDIn)
    'should be = ASF_Reserved_2 - 86D15241-311D-11D0-A3A4-00A0C90348F6
    If strRes2_GUID <> ASF_Reserved_2 Then
        eM$ = eM$ & vbCrLf & "** GUID Error in ASF_Reserved_2 = " & strRes2_GUID
    Else
        Debug.Print "Reserved GUID OK   = "; strRes2_GUID; "(ASF_Reserved_2)"
    End If
    Debug.Print "CodecEntriesCount  = "; .CodecEntriesCount; "("; Hex$(.CodecEntriesCount); ")"
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "CodecList", LastAddress$ & ": " & "ASF_Codec_List_Object", 3)
    Set tempNode = TV.Nodes.Add("CodecList", tvwChild, "strCodec_GUID", "GUID = " & strCodec_GUID, 4)
    Set tempNode = TV.Nodes.Add("CodecList", tvwChild, "Codec_QSize", "ObjectSize         = " & Codec_QSize & " (" & Hex$(Codec_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("CodecList", tvwChild, "strRes2_GUID", "strRes2_GUID       = " & strRes2_GUID, 4)
    Set tempNode = TV.Nodes.Add("CodecList", tvwChild, "CodecEntriesCount", "CodecEntriesCount  = " & .CodecEntriesCount, 4)
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Codec_List_Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objCodecListInfo"
GetCodecs
GoTo Exit_ParseCodecList
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseCodecList "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseCodecList"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseCodecList:
End Sub

Sub GetCodecs()
On Error GoTo Oops
Dim strInfo$
'
'Get #fNum, , objCodecListInfo
'If strGUID <> ASF_Codec_List_Object Then
'    eM$ =  "ASF_Codec_List_Object"
'    eM$ = eM$ & vbCrLf & "---------------------------"
'    eM$ = eM$ & vbCrLf & "** Error in strGUID = " & strGUID
'End If
'With objCodecListInfo
'    Debug.Print "-------------------ASF_Codec_List_Object-----------------------------"
'    Debug.Print "ObjectSize_Low   = "; .ObjectSize_Low; "("; Hex$(.ObjectSize_Low); ")"
'    Debug.Print "ObjectSize_High  = "; .ObjectSize_High; "("; Hex$(.ObjectSize_High); ")"
'    GUIDIn = .Reserved1
'    strGUID =  GUIDToString(GUIDIn)
'    Debug.Print "Reserved1 GUID = "; strGUID
'    Debug.Print "CodecEntriesCount  = "; .CodecEntriesCount; "("; Hex$(.CodecEntriesCount); ")"
'End With
'Debug.Print Hex$(Seek(fNum)); " After objCodecListInfo"
'
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before GetCodecs"
CodecCount = objCodecListInfo.CodecEntriesCount
ReDim Preserve Codec(CodecCount)
Debug.Print "Retrieving data for "; CodecCount; " Codec Entries."
eM$ = ""
'
'Field Name             Field Type  Size (bytes)
'----------------------------------------------
'Type                       WORD    2
'Codec Name Length          WORD    2
'Codec Name                 WCHAR  varies
'Codec Description Length   WORD    2
'Codec Description          WCHAR  varies
'Codec Information Length   WORD    2
'Codec Information          BYTE   varies
'
Set tempNode = TV.Nodes.Add("CodecList", tvwChild, "Codec", LastAddress$ & ": " & "Codec Entries", 3)
For i = MaxCTags To CodecCount - 1
    CTagIndex = i 'MaxCTags + i
    With Codec(CTagIndex)
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before Codec Data " & i
        Get #fNum, , .Type
        If .Type <> 2 Then
            MsgBox "This is not an audio codec!"
        End If
        Debug.Print "Codec Type = "; .Type
        'get Length of the name data
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before Codec Name " & i
        Get #fNum, , .NameLen
        Debug.Print "Codec NameLen = "; .NameLen 'namelen is 24 d
        ReDim .bytCodecName((2 * .NameLen) - 1) '(47) '48 characters '
        'get the name of the CTag
        Get #fNum, , .bytCodecName
        'convert the byte data to a string by forcing it into the string value
        .Name = .bytCodecName
        'trim the null from the Name
        .Name = Left$(.Name, Len(.Name) - 1)
        Debug.Print .Name
        '
        'get the Description length
        Get #fNum, , .DescrLen
        If .DescrLen > 0 Then
            ReDim .bytCodecDescr((2 * .DescrLen) - 1)
            'get the Descr of the CTag
            Get #fNum, , .bytCodecDescr
            'convert the byte data to a string by forcing it into the string value
            .Descr = .bytCodecDescr
            'trim the null from the Descr
            .Descr = Left$(.Descr, Len(.Descr) - 1)
            Debug.Print .Descr
        End If
        '
        'Get the Codec Data
        Get #fNum, , .DataLen
        If .DataLen > 0 Then
            ReDim .bytCodecData(.DataLen - 1)
            'get the Data of the CTag
            Get #fNum, , .bytCodecData
            'copy the data to our variant
            Debug.Print "Codec DataLen="; .DataLen
            CopyMemory .Data, ByVal VarPtr(.bytCodecData(0)), CLng(.DataLen)
            Debug.Print "(Format is "; GetFormatTag(CInt(.Data)); ")"
        End If
        '
        'parse the description
        '160 kbps, 44 kHz, stereo 1-pass CBR
        'cbr=Constant Bit Rate
        Dim aryCodec() As String
        aryCodec = Split(.Descr, ",")
    End With
    With InfoTag
        Select Case Left$(Codec(CTagIndex).Name, 21)
            Case "Windows Media Audio 9" '9.1"
                .Bitrate = 0
                Dim intKBPS As Integer
                Dim strKBPS As String
                Dim intLength As Long
                Dim intFoundFreq As Integer
                Dim intFoundMode As Integer
                intKBPS = InStr(1, UCase(aryCodec(0)), "KBPS")
                .Bitrate = 0
                If intKBPS > 0 Then
                    strKBPS = Trim(Mid(aryCodec(0), intKBPS - 4, 3))
                    If Not strKBPS = "" And IsNumeric(strKBPS) Then
                        .Bitrate = CLng(strKBPS) * 1000
                    End If
                End If
                If .Bitrate > 0 Then
                    intLength = CLng(((FileLen(AudioFileName) * 8) - 217480) / (CLng(.Bitrate) * 1000))
                    'if we can't get the length by the other way
                    '.Length = intLength
                    '.Duration = fTime(.Length)
                End If
                aryCodec(1) = Trim$(aryCodec(1))
                intFoundFreq = InStr(1, UCase(aryCodec(1)), "KHZ")
                Dim strFreq As String
                If intFoundFreq > 0 Then
                    strFreq = Left$(aryCodec(1), intFoundFreq - 1)
                    strFreq = Trim(Replace(strFreq, ",", ""))
                    If IsNumeric(strFreq) Then
                        .Frequency = CLng(strFreq & "000")
                    End If
                End If
'                If InStr(1, UCase(aryCodec(2)), "STEREO") > 0 Then
'                    .Mode = "Stereo"
'                End If
'                If InStr(1, UCase(aryCodec(2)), "MONO") > 0 Then
'                    .Mode = "Mono"
'                End If
                .Mode = aryCodec(2)
        End Select
    End With
    With Codec(CTagIndex)
        strInfo$ = .Name & " - " & .Descr & " = " & .Data & vbCrLf
        eM$ = eM$ & strInfo$
        Set tempNode = TV.Nodes.Add("Codec", tvwChild, ".Data" & CTagIndex, LastAddress$ & ": " & .Name, 3)
        Set tempNode = TV.Nodes.Add(".Data" & CTagIndex, tvwChild, "CDescr" & CTagIndex, "Descr = " & .Descr, 4)
        Set tempNode = TV.Nodes.Add(".Data" & CTagIndex, tvwChild, "CData" & CTagIndex, "Data  = " & .Data & "(Format is " & GetFormatTag(.Data) & ")", 4)
    End With
Next i
MaxCTags = MaxCTags + CodecCount
'
Debug.Print "----------------------------------------------------------"
Debug.Print eM$
Debug.Print "----------------------------------------------------------"
'
GoTo Exit_GetCodecs
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetCodecs "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GetCodecs"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetCodecs:
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After GetCodecs"
End Sub

Sub ParseStreamProperties()
On Error GoTo Oops
'-----------------------------ASF_Stream_Properties_Object----------------------------------
'we must first get ASF_Stream_Properties_Object
'this should be B7DC0791-A9B7-11CF-8EE6-00C00C205365
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objStreamInfo"
Get #fNum, , objStreamInfo
HeaderCompleteCount = HeaderCompleteCount + 1
GUIDIn = objStreamInfo.ObjectID
strStream_GUID = GUIDToString(GUIDIn)
Debug.Print "ObjectID = "; strStream_GUID
eM$ = "Executing " & ReturnGUIDName(strStream_GUID) & vbCrLf & strStream_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
'
If strStream_GUID <> ASF_Stream_Properties_Object Then
    eM$ = "ASF_Stream_Properties_Object"
    eM$ = eM$ & vbCrLf & "-------------------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strStream_GUID = " & vbCrLf & strStream_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
With objStreamInfo
    Debug.Print "-------------------ASF_Stream_Properties_Object-------------------------"
    Stream_QSize = QWORD(.ObjectSize) 'Valid values are at least 78 bytes
    Accum_HeaderQSize = Accum_HeaderQSize + Stream_QSize
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseStreamProperties"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    Debug.Print "ObjectSize       = "; Stream_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    GUIDIn = .StreamType
    'should be ASF_Audio_Media
    'F8699E40-5B4D-11CF-A8FD-00805F5C442B
    strStreamType_GUID = GUIDToString(GUIDIn)
    Debug.Print "strStreamType_GUID = "; strStreamType_GUID
    If strStreamType_GUID <> ASF_Audio_Media Then
        eM$ = "Error in ASF_Stream_Properties_Object StreamType"
        eM$ = eM$ & vbCrLf & "---------------------------"
        eM$ = eM$ & vbCrLf & "** Error in strStreamType_GUID = " & vbCrLf & strStreamType_GUID
        Debug.Print eM$
        MsgBox eM$, , "ASF_Stream_Properties_Object"
        eM$ = ""
    End If
    GUIDIn = .ErrorType
    'should be 'ASF_Audio_Spread
    'BFC3CD50-618F-11CF-8BB2-00AA00B4E220
    strError_GUID = GUIDToString(GUIDIn)
    Debug.Print "strError_GUID = "; strError_GUID
    If strError_GUID <> ASF_Audio_Spread Then
        eM$ = "Error in ASF_Stream_Properties_Object Error Type"
        eM$ = eM$ & vbCrLf & "** Error in strError_GUID = " & vbCrLf & strError_GUID
        eM$ = eM$ & vbCrLf & "at " & LastAddress$ '& Seek(fNum) & " (" & AddZero(Hex$(Seek(fNum)), 6) & "H)"
        Debug.Print eM$
        MsgBox eM$, , "ASF_Stream_Properties_Object"
        eM$ = ""
    End If
    '
    TimeOffsetQSize = QWORD(.TimeOffset)
    Debug.Print "TimeOffset     = "; TimeOffsetQSize; "("; Hex$(.TimeOffset.High); AddZero(Hex$(.TimeOffset.Low), 8); ")"
    Dim varTimeOffset As Variant
    varTimeOffset = CVar(TimeOffsetQSize * ((10 ^ -7)))
    Debug.Print "TimeOffset     = "; varTimeOffset; " sec."
    Debug.Print "TypeDataLen    = "; .TypeDataLen; "("; Hex$(.TypeDataLen); ")"
    ReDim bytTypeSpecific(.TypeDataLen - 1)
    Debug.Print "ErrorDataLen   = "; .ErrorDataLen; "("; Hex$(.ErrorDataLen); ")"
    ReDim bytErrorCorrection(.ErrorDataLen - 1)
    Debug.Print "Flags          = "; .Flags; "("; Hex$(.Flags); ")"
    If .Reserved <> 0 Then   'always 0
        'odd...it seems to be NEVER be 0
        'eM$ = eM$ & vbCrLf & "** Error in Reserved = " & .Reserved & " (" & Hex$(.Reserved) & ")"
    Else
        Debug.Print "Reserved OK?   = "; .Reserved; "("; Hex$(.Reserved); ") - always 0"
    End If
    fstrInfo = fstrInfo & AddSpace(.Reserved & " (" & Hex$(.Reserved) & ")", 10) & " | "
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Stream", LastAddress$ & ": " & "ASF_Stream_Properties_Object", 3)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "strStream_GUID", "GUID = " & strStream_GUID, 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "Stream_QSize", "ObjectSize     = " & Stream_QSize & " (" & Hex$(Stream_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "strStreamType_GUID", "strStreamType_GUID = " & strStreamType_GUID, 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "strError_GUID", "strError_GUID  = " & strError_GUID, 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "TimeOffset", "TimeOffset     = " & varTimeOffset & " sec.", 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "TypeDataLen", "TypeDataLen    = " & .TypeDataLen & " (" & Hex$(.TypeDataLen) & ")", 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "ErrorDataLen", "ErrorDataLen   = " & .ErrorDataLen & " (" & Hex$(.ErrorDataLen) & ")", 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "SP_Flags", "Flags          = " & .Flags, 4)
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "Reserved?", "Reserved       = " & .Reserved, 4)
End With
If Len(eM$) > 0 Then
    Debug.Print eM$
    MsgBox eM$, , "ASF_Stream_Properties_Object"
    eM$ = ""
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objStreamInfo"
 'Type-Specific Data
'Specifies type-specific format data.
' The structure for the Type-Specific Data field is determined by the
' value stored in the Stream Type field. The structure for the
' Type-Specific Data field for standard ASF media types is detailed in section 9.
'When the Stream Type of the Stream Properties Object has the value ASF_Audio_Media,
' the ASF audio media type that populates the Type-Specific Data field of the
' Stream Properties Object is represented using the following structure
' (the WAVEFORMATEX structure).
'Field name              Field type  Size (bytes)
'-----------------------------------------------------------------------------------
'Codec ID / Format Tag      WORD        2
'Number of Channels         WORD        2
'Samples Per Second         DWORD       4
'Avg # of Bytes Per Sec     DWORD       4
'Block Alignment            WORD        2
'Bits Per Sample            WORD        2
'Codec Specific Data Size   WORD        2
'Codec Specific Data        BYTE     varies
'
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before TypeSpecificData"
'we need to jump to the audio parser
If strStreamType_GUID = ASF_Audio_Media Then
    strAudio_GUID = ASF_Audio_Media
    ParseAudio
Else
    Get #fNum, , bytTypeSpecific
    'move the data to a string
    TypeSpecificData = bytTypeSpecific
    'TypeSpecificData = Left$(TypeSpecificData, Len(TypeSpecificData) - 1)
    Debug.Print "TypeSpecificData= "; TypeSpecificData
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After TypeSpecificData"
'
'Error Correction Data
'Specifies data specific to the error correction type.
' The structure for the Error Correction Data field is determined by
' the value stored in the Error Correction Type field.
' For example, an audio data stream might need to know how codec chunks
' were redistributed, or it might need a sample of encoded silence. For
' detailed information, see section 9.1.
Get #fNum, , ErrorCorrectionData 'bytErrorCorrection
'move the data to a string
'ErrorCorrectionData = bytErrorCorrection
'CopyMemory ByVal VarPtr(ErrorCorrectionData(0)), ByVal VarPtr(bytErrorCorrection(0)), CLng(objStreamInfo.ErrorDataLen)
ErrorCorrectionDataQSize = QWORD(ErrorCorrectionData)
Debug.Print "ErrorCorrectionData     = "; ErrorCorrectionDataQSize; "("; Hex$(ErrorCorrectionData.High); AddZero(Hex$(ErrorCorrectionData.Low), 8); ")"
    
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ErrorCorrectionData"
GoTo Exit_ParseStreamProperties
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseStreamProperties "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseStreamProperties"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
Debug.Assert (mError <> vbRetry)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseStreamProperties:
End Sub

Sub ParseAudio()
On Error GoTo Oops
'-----------------------------------ASF_Audio_Media--------------------------------------------
ASF_Audio_Media:
'ASF_Stream_Bitrate_Properties_Object
'then we dig out the ASF_Audio_Media
'F8699E40-5B4D-11CF-A8FD-00805F5C442B
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ASF_Audio_Media"
'
Get #fNum, , objAudioInfo 'ASF_Audio_Media
With objAudioInfo
    Debug.Print "-------------------ASF_Audio_Media-----------------------------"
    Debug.Print "CodecID        = "; .CodecID; "("; Hex$(.CodecID); ")"
    Debug.Print "(Format is "; GetFormatTag(.CodecID); ")"
    Debug.Print "NumChannels    = "; .NumChannels; "("; Hex$(.NumChannels); ")"
    Debug.Print "SamplesPerSec  = "; .SamplesPerSec; "("; Hex$(.SamplesPerSec); ")"
    Debug.Print "AvgBytesPerSec = "; .AvgBytesPerSec; "("; Hex$(.AvgBytesPerSec); ")"
    Debug.Print "BlockAlignment = "; .BlockAlignment; "("; Hex$(.BlockAlignment); ")"
    Debug.Print "BitsPerSample  = "; .BitsPerSample; "("; Hex$(.BitsPerSample); ")"
    Debug.Print "CodecDataSize  = "; .CodecDataSize; "("; Hex$(.CodecDataSize); ")"
    Debug.Print "CodecData"
    If .CodecDataSize > 0 Then
        ReDim bytCodecSpecificData(.CodecDataSize - 1)
        Get #fNum, , bytCodecSpecificData
        For i = 0 To .CodecDataSize - 1
            Debug.Print i; "= "; bytCodecSpecificData(i); "("; Hex$(bytCodecSpecificData(i)); ")"
        Next i
    End If
    InfoTag.Frequency = .SamplesPerSec
    InfoTag.Bitrate = ((.AvgBytesPerSec * (.BitsPerSample / .NumChannels)) \ 1000) * 1000
    Dim intLength As Long
    If InfoTag.Bitrate > 0 Then
        intLength = CLng(((FileLen(AudioFileName) * 8) - 217480) / (CLng(InfoTag.Bitrate) * 1000))
        'if we can't get the length by the other way
        '.Length = intLength
        '.Duration = fTime(.Length)
    End If
    Select Case .NumChannels
        Case Is = 1
            InfoTag.Mode = "Mono"
        Case Is = 2
            InfoTag.Mode = "Stereo"
        Case Is = 6
            InfoTag.Mode = "5.1 Surround"
    End Select
    '
    Set tempNode = TV.Nodes.Add("Stream", tvwChild, "Audio", LastAddress$ & ": " & "ASF_Audio_Media", 3)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "strAudio_GUID", "GUID = " & strAudio_GUID, 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "Format", GetFormatTag(.CodecID), 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "NumChannels", "NumChannels    = " & .NumChannels, 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "SamplesPerSec", "SamplesPerSec  = " & .SamplesPerSec & " (" & Hex$(.SamplesPerSec) & ")", 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "AvgBytesPerSec", "AvgBytesPerSec = " & .AvgBytesPerSec & " (" & Hex$(.AvgBytesPerSec) & ")", 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "BlockAlignment", "BlockAlignment = " & .BlockAlignment & " (" & Hex$(.BlockAlignment) & ")", 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "BitsPerSample", "BitsPerSample  = " & .BitsPerSample & " (" & Hex$(.BitsPerSample) & ")", 4)
    Set tempNode = TV.Nodes.Add("Audio", tvwChild, "CodecDataSize", "CodecDataSize  = " & .CodecDataSize & " (" & Hex$(.CodecDataSize) & ")", 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objAudioInfo"
GoTo Exit_ParseAudio
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseAudio "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseAudio"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseAudio:
End Sub

Sub ParseStreamBitrate()
On Error GoTo Oops
'--------------------------ASF_Stream_Bitrate_Properties_Object-----------------------------------------------------
'we must first get ASF_Stream_Bitrate_Properties_Object
'this should be 7BF875CE-468D-11D1-8D82-006097C9A2B2
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objBitRateInfo"
Get #fNum, , objBitRateInfo
HeaderCompleteCount = HeaderCompleteCount + 1
GUIDIn = objBitRateInfo.ObjectID
strBitRate_GUID = GUIDToString(GUIDIn)
Debug.Print "ObjectID = "; strBitRate_GUID
eM$ = "Executing " & ReturnGUIDName(strBitRate_GUID) & vbCrLf & strBitRate_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
'
If strBitRate_GUID <> ASF_Stream_Bitrate_Properties_Object Then
    eM$ = "ASF_Stream_Bitrate_Properties_Object"
    eM$ = eM$ & vbCrLf & "-------------------------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strBitRate_GUID = " & vbCrLf & strBitRate_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
With objBitRateInfo
    Debug.Print "-------------------ASF_Stream_BitRate_Properties_Object-------------------"
    BitRate_QSize = QWORD(.ObjectSize)
    Accum_HeaderQSize = Accum_HeaderQSize + BitRate_QSize
'   All of the files I have are laid out such that the Accumulated size at this point
'   matches the header object size.
    If Accum_HeaderQSize = ASF_QSize Then
        'MsgBox "Accum_HeaderQSize = ASF_QSize", , "ParseStreamBitrate"
    End If
    Accum_QSize = Accum_HeaderQSize
    frmMain.txtInfo.SelText = Accum_QSize & vbTab
    Debug.Print "ObjectSize       = "; BitRate_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    Debug.Print "BitrateRecords    = "; .BitrateRecords; "("; Hex$(.BitrateRecords); ")"
    ReDim BitRateEntry(.BitrateRecords)
    '
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "BitRate", LastAddress$ & ": " & "ASF_Stream_BitRate_Properties_Object", 3)
    Set tempNode = TV.Nodes.Add("BitRate", tvwChild, "strBitRate_GUID", "GUID = " & strBitRate_GUID, 4)
    Set tempNode = TV.Nodes.Add("BitRate", tvwChild, "BitRate_QSize", "ObjectSize  = " & BitRate_QSize & " (" & Hex$(BitRate_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("BitRate", tvwChild, "BitrateRecords", "RecordCount = " & .BitrateRecords, 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objBitRateInfo"
'
For i = 0 To objBitRateInfo.BitrateRecords - 1
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print LastAddress$; " Before BitrateRecord"
    Get #fNum, , BitRateEntry(i)
    With BitRateEntry(i)
        Debug.Print "Bitrate Flags = "; .Flags; "("; Hex$(.Flags); ")"
        Debug.Print "AvgBitrate    = "; .AvgBitrate; "("; Hex$(.AvgBitrate); ")"
        Set tempNode = TV.Nodes.Add("BitRate", tvwChild, "BitrateRecord" & i, LastAddress$ & ": " & "BitrateRecord" & i, 3)
        Set tempNode = TV.Nodes.Add("BitrateRecord" & i, tvwChild, "BitrateFlag" & i, "Bitrate Flags = " & .Flags & " (" & Hex$(.Flags) & ")", 4)
        Set tempNode = TV.Nodes.Add("BitrateRecord" & i, tvwChild, "AvgBitrate" & i, "AvgBitrate    = " & .AvgBitrate & " (" & Hex$(.AvgBitrate) & ")", 4)
    End With
Next i
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After All BitrateRecords"
GoTo Exit_ParseStreamBitrate
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseStreamBitrate "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseStreamBitrate"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseStreamBitrate:
End Sub

Sub ParseData()
On Error GoTo Oops
'-----------------------------------ASF_Data_Object--------------------------------------------
'the next one is
ASF_Data_Object:
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before ASF_Data_Object"
Get #fNum, , objDataInfo 'ASF_Data_Object
HeaderCompleteCount = HeaderCompleteCount + 1
GUIDIn = objDataInfo.ObjectID
'should be 75B22636-668E-11CF-A6D9-00AA0062CE6C
strData_GUID = GUIDToString(GUIDIn)
Debug.Print "GUID = "; strData_GUID
'
eM$ = "Executing " & ReturnGUIDName(strData_GUID) & vbCrLf & strData_GUID
Debug.Print eM$
If DebugMode = 1 Then
    MsgBox eM$
End If
eM$ = ""
'
If strData_GUID <> ASF_Data_Object Then
    eM$ = "ASF_Data_Object"
    eM$ = eM$ & vbCrLf & "-------------------------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strData_GUID = " & vbCrLf & strData_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
With objDataInfo
    Debug.Print "----------------------ASF_Data_Object-----------------------------"
    '
    'objDataInfo
    '    ObjectID As GuidStruct 'ASF_Data_Object
    '    ObjectSize As DlngSize 'at least 50 bytes
    '    FileID As GuidStruct
    '    TotalDataPackets As DlngSize 'must be equal to the Data Packet Count field in the File Properties Object
    '    Reserved As Integer. The value of this field shall set to 0x0101. (257D)
    'then the data packets
    Data_QSize = QWORD(.ObjectSize)
    Debug.Print "ObjectSize       = "; Data_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    TotalData_QSize = Data_QSize - Len(objDataInfo)
    Accum_QSize = Accum_QSize + Data_QSize
    frmMain.txtInfo.SelText = Accum_QSize & vbCrLf
    '
    GUIDIn = .FileID
    strDataFileID_GUID = GUIDToString(GUIDIn)
    Debug.Print "FileID GUID = "; strDataFileID_GUID
    'Specifies the unique identifier for this ASF file.
    'The value of this field shall be changed every time the file is modified in any way.
    'The value of this field shall be identical to the value of the
    'File ID field of the Header Object.
    If strDataFileID_GUID <> strFile_GUID Then
        eM$ = eM$ & vbCrLf & "*** Error in unique identifier for this ASF file *** "
        eM$ = eM$ & vbCrLf & "It must be identical to the value of the File ID field of the Header Object."
        eM$ = eM$ & vbCrLf & "strDataFileID_GUID    = " & strDataFileID_GUID
        eM$ = eM$ & vbCrLf & "FP strFile_GUID         = " & strFile_GUID
        eM$ = eM$ & vbCrLf
        eM$ = eM$ & vbCrLf & "Do you want to use the File ID field of the Header Object for both?"
        eM$ = eM$ & vbCrLf & "(Click Yes to use the Header Object File ID for both.)"
        eM$ = eM$ & vbCrLf & "(Click No to use the Data ASF File ID for both.)"
        eM$ = eM$ & vbCrLf & "(Click Cancel to leave both unchanged.)"
        eTitle$ = App.Title & ": Error in Subroutine ParseData "
        mError = MsgBox(eM$, vbYesNoCancel, eTitle$)
        If mError = vbYes Then
            strDataFileID_GUID = strFile_GUID
        ElseIf mError = vbNo Then
            strFile_GUID = strDataFileID_GUID
        End If
        eM$ = ""
    Else
        Debug.Print "strDataFileID_GUID  = "; strDataFileID_GUID
    End If
    '
    Packets_QSize = QWORD(.TotalDataPackets)
    'It must be equal to the Data Packet Count field in the File Properties Object
    If Packets_QSize <> DataPacketsQSize Then
        eM$ = eM$ & vbCrLf & "** Error in TotalDataPackets = " & Packets_QSize & " (" & Hex$(Packets_QSize) & ")"
        eM$ = eM$ & vbCrLf & "It must be equal to the Data Packet Count field in the File Properties Object."
    Else
        Debug.Print "TotalDataPackets  = "; Packets_QSize; "("; Hex$(.TotalDataPackets.High); AddZero(Hex$(.TotalDataPackets.Low), 8); ")"
    End If
    If .Reserved <> 257 Then 'always 257
        eM$ = eM$ & vbCrLf & "** Error in Reserved = " & .Reserved & " (" & Hex$(.Reserved) & ")"
    Else
        Debug.Print "Reserved OK       = "; .Reserved; "("; Hex$(.Reserved); ")"
    End If
    '
    Set tempNode = TV.Nodes.Add("Root", tvwChild, "Data", LastAddress$ & ": " & "ASF_Data_Object", 3)
    Set tempNode = TV.Nodes.Add("Data", tvwChild, "strData_GUID", "GUID = " & strData_GUID, 4)
    Set tempNode = TV.Nodes.Add("Data", tvwChild, "Data_QSize", "ObjectSize  = " & Data_QSize & " (" & Hex$(Data_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Data", tvwChild, "strDataFileID_GUID", "strDataFileID_GUID = " & strDataFileID_GUID, 4)
    Set tempNode = TV.Nodes.Add("Data", tvwChild, "Packets_QSize", "PacketSize  = " & Packets_QSize & " (" & Hex$(Packets_QSize) & ")", 4)
    Set tempNode = TV.Nodes.Add("Data", tvwChild, "DataReserved", "Reserved    = " & .Reserved, 4)
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After ASF_Data_Object"
HeaderEnd = Seek(fNum)
Debug.Print "HeaderEnd = "; HeaderEnd; " ("; LastAddress$; ")"
'ReDim StreamData(LOF(fNum) - HeaderEnd - 1)
ReDim StreamData(Data_QSize - Len(objDataInfo) - 1)
Get #fNum, HeaderEnd, StreamData  'this gets the whole rest of the file
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After Stream Data"
'
EndOfDataAddress = Seek(fNum) - 1
lngExcessData = LOF(fNum) - Accum_QSize
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Data_Object"
    frmMain.txtStatus.SelText = eM$ & vbCrLf
    eM$ = ""
End If
GoTo Exit_ParseData
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ParseData "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ParseData"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ParseData:
End Sub

Function ReturnGUIDName(GuidText As String) As String
Select Case GuidText
    Case Is = ASF_Header_Object
        ReturnGUIDName = "ASF_Header_Object"
    Case Is = ASF_Content_Description_Object
        ReturnGUIDName = "ASF_Content_Description_Object"
    Case Is = ASF_Extended_Content_Description_Object
        ReturnGUIDName = "ASF_Extended_Content_Description_Object"
    Case Is = ASF_File_Properties_Object
        ReturnGUIDName = "ASF_File_Properties_Object"
    Case Is = ASF_Header_Extension_Object
        ReturnGUIDName = "ASF_Header_Extension_Object"
    Case Is = ASF_Undocumented_Compatibility_GUID
        ReturnGUIDName = "ASF_Undocumented_Compatibility_GUID"
    Case Is = ASF_Metadata_Object
        ReturnGUIDName = "ASF_Metadata_Object"
    Case Is = ASF_Padding_Object
        ReturnGUIDName = "ASF_Padding_Object"
    Case Is = ASF_Extended_Stream_Properties_Object
        ReturnGUIDName = "ASF_Extended_Stream_Properties_Object"
    Case Is = ASF_Index_Placeholder_Object
        ReturnGUIDName = "ASF_Index_Placeholder_Object"
    Case Is = ASF_Metadata_Library_Object
        ReturnGUIDName = "ASF_Metadata_Library_Object"
    Case Is = ASF_Codec_List_Object
        ReturnGUIDName = "ASF_Codec_List_Object"
    Case Is = ASF_Stream_Properties_Object
        ReturnGUIDName = "ASF_Stream_Properties_Object"
    Case Is = ASF_Stream_Bitrate_Properties_Object
        ReturnGUIDName = "ASF_Stream_Bitrate_Properties_Object"
    Case Is = ASF_Data_Object
        ReturnGUIDName = "ASF_Data_Object"
    'below is not utilized
    Case Is = ASF_Audio_Media
        ReturnGUIDName = "ASF_Audio_Media"
    Case Is = ASF_TempFileHeader
        '75B22668-668E-11CF-A6D9-00AA0062CE6C
        ReturnGUIDName = "ASF_TempFileHeader"
    Case Else
        Debug.Print "************************************************"
        ReturnGUIDName = "Error!  We received another GUID"
        Debug.Print ReturnGUIDName
        Debug.Print "GUID = "; GuidText
        Debug.Print "************************************************"
End Select
End Function
'=========================================================================================

Function Get_WMA_Header(ByVal strFileName As String)
'wma Header signature
'30 26 B2 75 8E 66 CF 11 A6 D9 00 AA 00 62 CE 6C
On Error GoTo Oops
Dim strData As String
Dim strData2 As String
Dim strSDKVer As String
Dim intGoto As Integer
Dim i As Integer
Dim intStart As String
Dim intStop As String
Dim intTitleLen As Integer
Dim intArtistLen As String
Dim intCopyLen As Integer
Dim strData3 As String
Dim intTempData As String
Dim myFSO As New FileSystemObject
Dim myTextStream As TextStream
'
Set myTextStream = myFSO.OpenTextFile(strFileName, ForReading, False)
fstrFullHeader = myTextStream.Read(Val("&H3000"))
myTextStream.Close
're-open
Set myTextStream = myFSO.OpenTextFile(strFileName, ForReading, False)
intGoto = 5000
myTextStream.Skip (10)
DoAgain:
For i = 1 To intGoto
    strData = myTextStream.Read(1)
    If Not Asc(strData) = 0 Then
        strData2 = strData2 & Chr(CLng(Asc(strData)))
    Else
        i = i - 1
    End If
    'myTextStream.Skip (1)
Next i
myTextStream.Close
Set myFSO = Nothing
Set myTextStream = Nothing
fstrHeader = strData2
InfoTag.FileSize = FileLen(strFileName)
intStart = InStr(1, strData2, "3&²uŽfÏ¦ÙªbÎl") + 15
'tp above string is the beginning of the wma file, but it should be 0 as the first
'character, not 3.  I think that the reason it works is because the string that is being
'built above skips 0's.  I don't understand why it shows up after the beginning.
'and why do we skip the first 10 characters?
'The string to identify the wma header actually appears twice in the first 60 bytes.
'below is each byte in hex.  The bytes above ^^ indicate bytes that vary from file to file
'
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
'40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
'This is then followed by some indicators to determine the next data.
'CE 04 00 00  00 00 00 00  13 00 12 00   precedes WM/Track
'DC 04 00 00  00 00 00 00  14 00 14 00   precedes WM/Lyrics
'00 10 00 is before WM/Year....
''it also appears that the tags are not always in the same order.
'
If intStart > 0 Then
    intStop = intStart + InStr(1, Mid(strData2, intStart + 1, Len(strData2) - intStart - 1), "@¤ÐÒ")
End If
If intStart > 0 And intStop > 0 And intStop > intStart Then
    strData3 = Mid(strData2, intStart, intStop - intStart)
    Dim intReadDat As Integer
    intReadDat = 1
    If Mid(strData3, 1, 1) = Chr(1) Or Mid(strData3, 1, 1) = Chr(2) Then
        intReadDat = 2
    End If
    intTitleLen = (Asc(Mid(strData3, intReadDat, 1)) / 2) - 1
    intArtistLen = (Asc(Mid(strData3, intReadDat + 1, 1)) / 2) - 1
    intCopyLen = (Asc(Mid(strData3, intReadDat + 2, 1)) / 2) - 1
    If intTitleLen + intArtistLen = Len(strData3) - 2 Then
        'artist and title
        InfoTag.Artist = Mid(strData3, intReadDat + 2 + intTitleLen, intArtistLen)
        InfoTag.Title = Mid(strData3, intReadDat + 2, intTitleLen)
    End If
    If intTitleLen + intArtistLen + intCopyLen = Len(strData3) - 2 - intReadDat Then
        InfoTag.Artist = Mid(strData3, intReadDat + 3 + intTitleLen, intArtistLen)
        InfoTag.Title = Mid(strData3, intReadDat + 3, intTitleLen)
        InfoTag.IsCopyright = True
    End If
End If
Dim intKBPS As Integer
Dim strKBPS As String
Dim intLength As Long
Dim intFoundAl As Integer
Dim intFoundFreq As Integer
Dim intFoundMode As Integer
intKBPS = InStr(1, UCase(strData2), "KBPS")
InfoTag.Bitrate = 0
If intKBPS > 0 Then
    strKBPS = Trim(Mid(strData2, intKBPS - 4, 3))
    If Not strKBPS = "" And IsNumeric(strKBPS) Then
        InfoTag.Bitrate = CInt(strKBPS)
    End If
End If
If InfoTag.Bitrate > 0 Then
    intLength = CLng(((FileLen(strFileName) * 8) - 217480) / (CLng(InfoTag.Bitrate) * 1000))
    InfoTag.Length = intLength
    InfoTag.Duration = fTime(InfoTag.Length)
End If
intFoundFreq = InStr(1, UCase(strData2), "KHZ")
Dim strFreq As String
If intFoundFreq > 0 Then
    strFreq = Mid(strData2, intFoundFreq - 5, 4)
    strFreq = Trim(Replace(strFreq, ",", ""))
    If IsNumeric(strFreq) Then
        InfoTag.Frequency = CLng(strFreq & "000")
    End If
End If
If InStr(1, UCase(strData2), "STEREO") > 0 Then
    InfoTag.Mode = "Stereo"
End If
If InStr(1, UCase(strData2), "MONO") > 0 Then
    InfoTag.Mode = "Mono"
End If
If InStr(1, UCase(strData2), "SIGNATURE") > 0 Then
    InfoTag.IsLicensed = True
End If
Dim intFoundYear As Integer
intFoundYear = InStr(1, UCase(strData2), "WM/YEAR")
If intFoundYear > 0 Then
    intFoundYear = intFoundYear + 7
    InfoTag.Year = Mid(strData2, intFoundYear + 1, CInt(Asc(Mid(strData2, intFoundYear, 1)) / 2) - 1)
End If
intFoundAl = InStr(1, UCase(strData2), "WM/ALBUMTITLE")
If intFoundAl > 0 Then
    intFoundAl = intFoundAl + 13
    InfoTag.Album = Mid(strData2, intFoundAl + 1, CInt(Asc(Mid(strData2, intFoundAl, 1)) / 2) - 1)
End If
'below added by tp
Dim intFoundComp As Integer
intFoundComp = InStr(1, UCase(strData2), "WM/COMPOSER")
If intFoundComp > 0 Then
    intFoundComp = intFoundComp + 11
    InfoTag.Composer = Mid(strData2, intFoundComp + 1, CInt(Asc(Mid(strData2, intFoundComp, 1)) / 2) - 1)
End If
Dim intFoundPub As Integer
intFoundPub = InStr(1, UCase(strData2), "WM/PUBLISHER")
If intFoundPub > 0 Then
    intFoundPub = intFoundPub + 12
    InfoTag.Publisher = Mid(strData2, intFoundPub + 1, CInt(Asc(Mid(strData2, intFoundPub, 1)) / 2) - 1)
End If
Dim intFoundGenre As Integer
intFoundGenre = InStr(1, UCase(strData2), "WM/GENRE")
If intFoundGenre > 0 Then
    intFoundGenre = intFoundGenre + 8
    InfoTag.Genre = Mid(strData2, intFoundGenre + 1, CInt(Asc(Mid(strData2, intFoundGenre, 1)) / 2) - 1)
End If
Dim intFoundTrackNum As Integer
intFoundTrackNum = InStr(1, UCase(strData2), "WM/TRACKNUMBER")
If intFoundTrackNum > 0 Then
    intFoundTrackNum = intFoundTrackNum + 14
    Dim strTrack As String
    'fbyteTrackNumber
    strTrack = Mid(strData2, intFoundTrackNum + 2, CInt(Asc(Mid(strData2, intFoundTrackNum, 1)) / 2) - 1)
    InfoTag.TrackNumber = Val(Asc(strTrack))
    Debug.Print "strTrack = "; Asc(strTrack), strTrack
End If
'WM/COMPOSER"NATALIE MERCHANT
'WM/PUBLISHERELEKTRA
'WM/GENRE ROCK
'WM/TrackNumber
GoTo Exit_Get_WMA_Header
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Get_WMA_Header "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Get_WMA_Header"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Get_WMA_Header:
End Function

Function fTime(lngLength As Long) As String
Dim intRemMin As Integer
Dim strRemMin As String
Dim strMin As String
Dim intMin As Integer
Dim intSec As Integer
Dim strHour As String
Dim intHour As Integer
strHour = CStr((lngLength \ 3600)) 'was 360
If InStr(1, strHour, ".") > 0 Then
    strHour = Left(strHour, InStr(1, strHour, ".") - 1)
End If
If Not strHour = "" Then
    intHour = CInt(strHour)
Else
    intHour = 0
End If
strRemMin = CStr(CDbl(lngLength Mod 3600)) 'was 360
If InStr(1, strRemMin, ".") > 0 Then
    strRemMin = Left(strRemMin, InStr(1, strRemMin, ".") - 1)
End If
If Not strRemMin = "" Then
    intRemMin = CInt(strRemMin)
Else
    intRemMin = 0
End If
strMin = CStr((intRemMin \ 60))
If InStr(1, strMin, ".") > 0 Then
    strMin = Left(strMin, InStr(1, strMin, ".") - 1)
End If
If Not strMin = "" Then
    intMin = CInt(strMin)
Else
    intMin = 0
End If
intSec = CInt(lngLength Mod 60)
If intHour > 0 Then
    fTime = Double_digits(intHour) & ":" & Double_digits(intMin) & ":" & Double_digits(intSec)
ElseIf intMin > 9 Then
    fTime = Double_digits(intMin) & ":" & Double_digits(intSec)
Else
    fTime = Str$(intMin) & ":" & Double_digits(intSec)
End If
End Function

Function Double_digits(intInput As Integer) As String
Dim strInput As String
strInput = CStr(intInput)
If Len(strInput) < 2 Then
    strInput = "0" & strInput
End If
Double_digits = strInput
End Function

Sub RearrangeFile()
'most of the files I have are in the following order
'----------------------------File Structure----------------------------
' 0  => ASF_Header_Object
' 1  => ASF_Content_Description_Object
' 2  => ASF_Extended_Content_Description_Object
' 3  => ASF_File_Properties_Object
' 4  => ASF_Header_Extension_Object
' 5  => ASF_Undocumented_Compatibility_GUID
' 6  => ASF_Metadata_Object
' 7  => ASF_Padding_Object
' 8  => ASF_Extended_Stream_Properties_Object
' 9  => ASF_Index_Placeholder_Object
' 10  => ASF_Metadata_Library_Object
' 11  => ASF_Codec_List_Object
' 12  => ASF_Stream_Properties_Object
' 13  => ASF_Stream_Bitrate_Properties_Object
' 14  => ASF_Data_Object
'----------------------------------------------------------------------
'the temp files have the following structure
' 0  => ASF_TempFileHeader
' 1  => ASF_File_Properties_Object
' 2  => ASF_Header_Extension_Object
' 3  => ASF_Undocumented_Compatibility_GUID
' 4  => ASF_Metadata_Object
' 5  => ASF_Padding_Object
' 6  => ASF_Extended_Stream_Properties_Object
' 7  => ASF_Index_Placeholder_Object
' 8  => ASF_Extended_Content_Description_Object
' 9  => ASF_Codec_List_Object
' 10  => ASF_Stream_Properties_Object
' 11  => ASF_Stream_Bitrate_Properties_Object
' 12  => ASF_Content_Description_Object
' 13  => ASF_Data_Object
'----------------------------------------------------------------------
're-assemble the data?
ObjectStack(0) = "ASF_Header_Object"
ObjectStack(1) = "ASF_Content_Description_Object"
ObjectStack(2) = "ASF_Extended_Content_Description_Object"
ObjectStack(3) = "ASF_File_Properties_Object"
ObjectStack(4) = "ASF_Header_Extension_Object"
ObjectStack(5) = "ASF_Undocumented_Compatibility_GUID"
ObjectStack(6) = "ASF_Metadata_Object"
ObjectStack(7) = "ASF_Padding_Object"
ObjectStack(8) = "ASF_Extended_Stream_Properties_Object"
ObjectStack(9) = "ASF_Index_Placeholder_Object"
ObjectStack(10) = "ASF_Metadata_Library_Object"
ObjectStack(11) = "ASF_Codec_List_Object"
ObjectStack(12) = "ASF_Stream_Properties_Object"
ObjectStack(13) = "ASF_Stream_Bitrate_Properties_Object"
ObjectStack(14) = "ASF_Data_Object"
StackNum = 15
End Sub

Sub SaveWMAHeader(fName As String)
On Error GoTo Oops
Dim i As Integer
Dim j As Integer
Dim fNum As Integer
Dim eM$
'
Dim WatchDog#
WatchDog# = Timer
HeaderCompleteCount = 0
SubHeaderCompleteCount = 0
If Len(Dir(fName)) > 0 Then
    Kill fName
End If
fNum = FreeFile
Open fName For Binary As fNum
Dim ObjNum As Integer
RearrangeFile
'loop through the objects and build the file exactly as it was before
For ObjNum = 0 To StackNum - 1
    Select Case ObjectStack(ObjNum)
        Case Is = "ASF_TempFileHeader", "ASF_Header_Object"
            PutMainHeader
        Case Is = "ASF_Content_Description_Object"
            PutContentHeader
        Case Is = "ASF_Extended_Content_Description_Object"
            PutExtContentData
        Case Is = "ASF_File_Properties_Object"
            PutFileProperties
        '
        Case Is = "ASF_Header_Extension_Object"
            PutHeaderExtension
        '
        '--------------------------------------------------
        'Below data is all part of the extended header
        'they don't count as header objects
        '
        Case Is = "ASF_Undocumented_Compatibility_GUID"
            PutCompatibility
        Case Is = "ASF_Metadata_Object"
            PutMetaData
        Case Is = "ASF_Padding_Object"
            PutPadding
        Case Is = "ASF_Extended_Stream_Properties_Object"
            PutExtendedStream
        Case Is = "ASF_Index_Placeholder_Object"
            PutPlaceholder
        Case Is = "ASF_Metadata_Library_Object"
            PutMetadataLibrary
        '--------------------------------------------------
        Case Is = "ASF_Codec_List_Object"
            PutCodecList
        Case Is = "ASF_Stream_Properties_Object"
            PutStreamProperties
        Case Is = "ASF_Stream_Bitrate_Properties_Object"
            PutStreamBitrate
        Case Is = "ASF_Data_Object"
            PutData
            'Next comes the data packets
            'An ASF Data Packet is a conveniently sized grouping of complete or
            'fragmented media objects from several digital media streams.
            '
        Case Else
            eM$ = "Unhandled GUID!" & vbCrLf
            eM$ = eM$ & ReturnGUIDName(strNextGUID) & vbCrLf
            eM$ = eM$ & strNextGUID
            MsgBox eM$, vbCritical, "Unhandled GUID!"
            eM$ = ""
            WatchDog# = Timer
            Exit For
    End Select
    If IsInIDE = False Then
        If Timer - WatchDog# > 60 Then
            eM$ = "WatchDog TimeOut!" & vbCrLf
            MsgBox eM$, vbCritical, "Error in SaveWMAHeader"
            eM$ = ""
            Exit For
        End If
    End If
Next ObjNum
'
'PutMainHeader
''
'PutContentHeader
''
'PutExtContentData
''
'PutFileProperties
''
'PutHeaderExtension
''
''--------------------------------------------------
''Below data is all part of the extended header
''they don't count as header objects
''
'PutCompatibility
'PutMetaData
'PutPadding
'PutExtendedStream
'PutPlaceholder
'PutMetadataLibrary
''--------------------------------------------------
'PutCodecList
''
'PutStreamProperties
''
'PutStreamBitrate
''
'PutData
'Next comes the data packets
'An ASF Data Packet is a conveniently sized grouping of complete or
'fragmented media objects from several digital media streams.
'
AllDone:
Close fNum
'now try to re-open and test
'ReadWMAHeader (fName)
'
GoTo Exit_PutDescriptor
Oops:
'Abort=3,Retry=4,IeTitle$ = App.Title & ": Error in Subroutine PutDescriptor "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutDescriptor"
eMess$ = eMess$ & IIf(Erl = 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutDescriptor:
End Sub

'=========================================================================================
'In order to write the file, we need to assemble the objects

Sub PutMainHeader()
Debug.Print "---------------------------"
Debug.Print "Saving objMainHeader..."
Dim ObjNum As Integer
'loop through the objects and build the size for the main header
For ObjNum = 0 To StackNum - 1
    Select Case ObjectStack(ObjNum)
        Case Is = "ASF_TempFileHeader", "ASF_Header_Object"
            'first initialize the write size to the size of the main header
            AccumWrite_QSize = Len(objMainHeader)
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Content_Description_Object"
            'Then build the content header and data to get the proper sizes
            BuildContentHeader
            AccumWrite_QSize = AccumWrite_QSize + CD_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Extended_Content_Description_Object"
            BuildExtContentData
            AccumWrite_QSize = AccumWrite_QSize + ExtContSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_File_Properties_Object"
            'save the file properties for the end so we can calculate the file size
            'BuildFileProperties
            'no we should calculate it here to avoid problems in debugging
            'calculate the file properties object size
            FP_QSize = Len(objFileInfo)
            AccumWrite_QSize = AccumWrite_QSize + FP_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Header_Extension_Object"
            'I'm running BuildHeaderExtension twice...
            'don't add the extended data size to the accum. value as it encompasses several other objects
            'add the len of the header extension object instad
            AccumWrite_QSize = AccumWrite_QSize + Len(objHeaderExtInfo)
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
            '
            '--------------------------------------------------
            'Below data is all part of the extended header
            'they don't count as header objects
        Case Is = "ASF_Undocumented_Compatibility_GUID"
            BuildCompatibility
            AccumWrite_QSize = AccumWrite_QSize + Comp_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Metadata_Object"
            'calculate the metadata
            BuildMetaData
            AccumWrite_QSize = AccumWrite_QSize + MetaData_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Padding_Object"
            BuildPadding
            AccumWrite_QSize = AccumWrite_QSize + Pad_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Extended_Stream_Properties_Object"
            BuildExtendedStream
            AccumWrite_QSize = AccumWrite_QSize + ExtStream_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Index_Placeholder_Object"
            'do the placeholder
            BuildPlaceholder
            AccumWrite_QSize = AccumWrite_QSize + PH_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Metadata_Library_Object"
            BuildMetaDataLibrary
            AccumWrite_QSize = AccumWrite_QSize + MetaDataLib_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
            '--------------------------------------------------
        Case Is = "ASF_Codec_List_Object"
            BuildCodecs
            AccumWrite_QSize = AccumWrite_QSize + Codec_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Stream_Properties_Object"
            BuildStreamProperties
            AccumWrite_QSize = AccumWrite_QSize + Stream_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Stream_Bitrate_Properties_Object"
            BuildStreamBitrate
            AccumWrite_QSize = AccumWrite_QSize + BitRate_QSize
            frmMain.txtInfo.SelText = AccumWrite_QSize & vbTab
        Case Is = "ASF_Data_Object"
            'Next comes the data packets
            'An ASF Data Packet is a conveniently sized grouping of complete or
            'fragmented media objects from several digital media streams.
            '
        Case Else
            eM$ = "Unhandled GUID!" & vbCrLf
            eM$ = eM$ & ReturnGUIDName(strNextGUID) & vbCrLf
            eM$ = eM$ & strNextGUID
            MsgBox eM$, vbCritical, "Unhandled GUID!"
            eM$ = ""
            Exit For
    End Select
Next ObjNum
'
'we have to build the extension here since it relies on above
BuildHeaderExtension
'
'Calculate the size for use in the File Properties
calcFileQSize = Len(objMainHeader)
calcFileQSize = calcFileQSize + QWORD(objCD_Info.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objExtContInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objFileInfo.ObjectSize)
'don't add the Header Extension object size to the accum. value as it
'encompasses several other objects that are already counted.
'Add the length of the object instead.
'Debug.Print QWORD(objHeaderExtInfo.ObjectSize) - Len(objHeaderExtInfo)
'calcFileQSize = calcFileQSize + QWORD(objHeaderExtInfo.ObjectSize)
calcFileQSize = calcFileQSize + Len(objHeaderExtInfo)
calcFileQSize = calcFileQSize + QWORD(objCompatibilityInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objMetaDataInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objPaddingInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objExtStreamInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objPlaceholderInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objMetaDataLibraryInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objCodecListInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objStreamInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objBitRateInfo.ObjectSize)
calcFileQSize = calcFileQSize + QWORD(objDataInfo.ObjectSize)
eM$ = eM$ & vbCrLf & "Actual FileSize         = " & vbTab & InfoTag.FileSize & " (" & Hex$(InfoTag.FileSize) & ")"
eM$ = eM$ & vbCrLf & "Calculated FileQSize   = " & vbTab & calcFileQSize & " (" & Hex$(calcFileQSize) & ")"
eM$ = eM$ & vbCrLf & "Difference              = " & vbTab & InfoTag.FileSize - calcFileQSize & " (" & Hex$(InfoTag.FileSize - calcFileQSize) & ")"
eM$ = eM$ & vbCrLf & "Accum Read FileSize  = " & vbTab & AccumRead_QSize & vbTab & " (" & Hex$(AccumRead_QSize) & ")"
eM$ = eM$ & vbCrLf & "Accum Read-FileQSize = " & vbTab & AccumRead_QSize - calcFileQSize & vbTab & " (" & Hex$(AccumRead_QSize - calcFileQSize) & ")"
eM$ = eM$ & vbCrLf & "Excess Data Count = " & vbTab & lngExcessData & " (" & Hex$(lngExcessData) & ")"
Debug.Print eM$
frmMain.txtStatus.SelText = eM$ & vbCrLf
eM$ = ""
If calcFileQSize <> InfoTag.FileSize Then
    eM$ = "*** Error in FileSize ***" & vbCrLf & eM$
    eM$ = eM$ & vbCrLf & "Actual FileSize        = " & InfoTag.FileSize & vbTab & " (" & Hex$(InfoTag.FileSize) & ")"
    eM$ = eM$ & vbCrLf & "FP Object FileSize     = " & calcFileQSize & vbTab & " (" & Hex$(calcFileQSize) & ")"
    'MsgBox eM$, , "ASF_File_Properties_Object"
    frmMain.txtStatus.SelText = eM$ & vbCrLf
    frmMain.txtStatus.SelText = "Using calculated filesize - " & calcFileQSize & vbCrLf
End If
eM$ = ""
BuildFileProperties
'
'Now we can write the new header
WriteHeader:
With objMainHeader
    .ObjectID = GUIDStrToGUIDStruct(ASF_Header_Object)
    'Debug.Print ReturnGUIDName(GUIDToString(.ObjectID))
    .ObjectSize.High = HiDWORD(AccumWrite_QSize)
    .ObjectSize.Low = LoDWORD(AccumWrite_QSize)
    .HeaderObjects = 7
    .Reserved1 = 1 'always 1
    .Reserved2 = 2 'always 2
End With
Put #fNum, , objMainHeader
End Sub

Sub BuildContentHeader()
On Error GoTo Oops
ReDim TempString(5)
Debug.Print "---------------------------"
Debug.Print "Building ContentHeader..."
With objCD_Info
    .ObjectID = GUIDStrToGUIDStruct(ASF_Content_Description_Object)
    'object size is calculated below
    If Len(InfoTag.Title) > 0 Then
        TempString(0) = StrConv(InfoTag.Title & Chr$(0), vbUnicode)
        .TitleLen = Len(TempString(0))
        'ReDim aryTitle(.TitleLen)
        aryTitle = StrConv(TempString(0), vbFromUnicode)
    End If
    If Len(InfoTag.Artist) > 0 Then
        TempString(1) = StrConv(InfoTag.Artist & Chr$(0), vbUnicode)
        .ArtistLen = Len(TempString(1))
        aryArtist = StrConv(TempString(1), vbFromUnicode)
    End If
    If Len(InfoTag.Copyright) > 0 Then
        TempString(2) = StrConv(InfoTag.Copyright & Chr$(0), vbUnicode)
        .CopyrightLen = Len(TempString(2))
        aryCopyright = StrConv(TempString(2), vbFromUnicode)
    End If
    If Len(InfoTag.Description) > 0 Then
        TempString(3) = StrConv(InfoTag.Description & Chr$(0), vbUnicode)
        .DescriptionLen = Len(TempString(3))
        aryDescription = StrConv(TempString(3), vbFromUnicode)
    End If
    If Len(InfoTag.Rating) > 0 Then
        TempString(4) = StrConv(InfoTag.Rating & Chr$(0), vbUnicode)
        .RatingLen = Len(TempString(4))
        aryRating = StrConv(TempString(4), vbFromUnicode)
    End If
    ' 34 bytes + size of content
    CD_QSize = Len(objCD_Info) + .TitleLen + .ArtistLen + .CopyrightLen + .DescriptionLen + .RatingLen
    .ObjectSize.High = HiDWORD(CD_QSize)
    .ObjectSize.Low = LoDWORD(CD_QSize)
End With
GoTo Exit_BuildContentHeader
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine BuildContentHeader "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in BuildContentHeader"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_BuildContentHeader:
End Sub

Sub PutContentHeader()
On Error GoTo Oops
Debug.Print "---------------------------"
Debug.Print "Saving objCD_Info..."
'
Put #fNum, , objCD_Info
'now set the len for the biblio tags and store the data
Put #fNum, , aryTitle() 'this Puts the block of bytes for the title
Put #fNum, , aryArtist() 'this Puts the block of bytes for the Artist
Put #fNum, , aryCopyright() 'this Puts the block of bytes for the copyright
Put #fNum, , aryDescription() 'this Puts the block of bytes for the Description
Put #fNum, , aryRating() 'this Puts the block of bytes for the Rating
'
HeaderCompleteCount = HeaderCompleteCount + 1
GoTo Exit_PutContentHeader
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutContentHeader "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutContentHeader"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutContentHeader:
End Sub

Sub BuildExtContentData()
On Error GoTo Oops
Dim tmpLong As Long
Dim tmpInt As Integer
Dim tmpSingle As Single
Dim tmpDouble As Double
Dim dLen As Long
eM$ = ""
Debug.Print "---------------------------"
Debug.Print "Building objExtContInfo..."
'
If DescriptorCount = 3 Then
    DescriptorCount = 15
    MaxTags = 15
    ReDim Preserve Descriptor(DescriptorCount)
    'lets add some more extended content....
    TagIndex = 0
    With Descriptor(TagIndex)
        '"WM/Track =  0 "
        .NameLen = 18   '(12)
        .Name = "WM/Track"
        'VB Data Type =  3  => Long
        .DataType = 3
        .DataLen = 4   '(4)
        .Data = 0
        '-------------End of  0 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/Lyrics = "
        .NameLen = 20   '(14)
        .Name = "WM/Lyrics"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 2   '(2)
        .Data = ""
        '-------------End of  1 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/MediaPrimaryClassID = {D1607DBC-E323-4BE2-86A1-48A42A28441E}"
        .NameLen = 46   '(2E)
        .Name = "WM/MediaPrimaryClassID"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 78   '(4E)
        .Data = "{D1607DBC-E323-4BE2-86A1-48A42A28441E}"
        '-------------End of  2 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WMFSDKVersion = 10.00.00.3646"
        .NameLen = 28   '(1C)
        .Name = "WMFSDKVersion"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 28   '(1C)
        .Data = "10.00.00.3646"
        '-------------End of  3 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WMFSDKNeeded = 0.0.0.0000"
        .NameLen = 26   '(1A)
        .Name = "WMFSDKNeeded"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 22   '(16)
        .Data = "0.0.0.0000"
        '-------------End of  4 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"IsVBR = False"
        .NameLen = 12   '(C)
        .Name = "IsVBR"
        'VB Data Type =  11  => Boolean
        .DataType = 2
        .DataLen = 4   '(4)
        .Data = "False"
        '-------------End of  5 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/TrackNumber =  1 "
        .NameLen = 30   '(1E)
        .Name = "WM/TrackNumber"
        'VB Data Type =  3  => Long
        .DataType = 3
        .DataLen = 4   '(4)
        .Data = 1
        '-------------End of  6 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/Year = unk"
        .NameLen = 16   '(10)
        .Name = "WM/Year"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 10   '(A)
        .Data = "Unknown"
        '-------------End of  7 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/Composer = Poe"
        .NameLen = 24   '(18)
        .Name = "WM/Composer"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 8   '(8)
        .Data = "Unknown"
        '-------------End of  8 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/Publisher = Atlantic"
        .NameLen = 26   '(1A)
        .Name = "WM/Publisher"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 18   '(12)
        .Data = "Unknown"
        '-------------End of  9 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/Genre = Rock"
        .NameLen = 18   '(12)
        .Name = "WM/Genre"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 10   '(A)
        .Data = "Rock"
        '-------------End of  10 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/AlbumTitle = Haunted"
        .NameLen = 28   '(1C)
        .Name = "WM/AlbumTitle"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 16   '(10)
        .Data = "Unknown"
        '-------------End of  11 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/AlbumArtist = Poe"
        .NameLen = 30   '(1E)
        .Name = "WM/AlbumArtist"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 8   '(8)
        .Data = "Unknown"
        '-------------End of  12 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/Provider = AMG"
        .NameLen = 24   '(18)
        .Name = "WM/Provider"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 8   '(8)
        .Data = "Unknown"
        '-------------End of  13 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/ProviderRating = 6"
        .NameLen = 36   '(24)
        .Name = "WM/ProviderRating"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 4   '(4)
        .Data = "Unknown"
        '-------------End of  14 --------------------
    End With
    TagIndex = TagIndex + 1
    With Descriptor(TagIndex)
        '"WM/ProviderStyle = Rock"
        .NameLen = 34   '(22)
        .Name = "WM/ProviderStyle"
        'VB Data Type =  8  => String
        .DataType = 0
        .DataLen = 10   '(A)
        .Data = "Rock"
        ExtContSize = 1204  '(4B4)
        '-------------End of  15 --------------------
    End With
End If
'
ReDim TempString(DescriptorCount)
Erase VarOutput()
ExtContSize = 0
For i = 0 To DescriptorCount - 1
    With InfoTag
        Select Case Descriptor(i).Name
            Case "WM/Track"
                Descriptor(i).Data = .TrackNumber - 1
            Case "WM/Lyrics"
            Case "WM/MediaPrimaryClassID"
            Case "WMFSDKVersion"
            Case "WMFSDKNeeded"
            Case "IsVBR"
            Case "WM/TrackNumber"
                Descriptor(i).Data = .TrackNumber
            Case "WM/Year"
                Descriptor(i).Data = .Year
            Case "WM/EncodingTime"
            Case "WM/UniqueFileIdentifier"
            Case "WM/Composer"
                Descriptor(i).Data = .Composer
            Case "WM/Publisher"
                Descriptor(i).Data = .Publisher
            Case "WM/Genre"
                Descriptor(i).Data = .Genre
            Case "WM/AlbumTitle"
                Descriptor(i).Data = .Album
            Case "WM/AlbumArtist"
                Descriptor(i).Data = .AlbumArtist
            Case "WM/MCDI"
            Case "WM/Provider"
            Case "WM/ProviderRating"
                Descriptor(i).Data = .ProviderRating
            Case "WM/ProviderStyle"
        End Select
    End With
    With Descriptor(i)
        TempString(i) = StrConv(.Name & Chr$(0), vbUnicode)
        .NameLen = Len(TempString(i))
        .bytTagName = StrConv(TempString(i), vbFromUnicode)
        ExtContSize = ExtContSize + LenB(.NameLen)
        ExtContSize = ExtContSize + .NameLen
        '
        'You can determine how the data in a Variant is treated using the
        'VarType function or TypeName function From VarType
        ReDim Preserve VarOutput(i) As Variant
        ReDim .bytTagData(0 To .DataLen - 1)
        Select Case VarType(.Data)
            'Case Is = 0 'vbEmpty 0 Empty (uninitialized)
            'Case Is = 1 'vbNull 1 Null (no valid data)
            Case Is = 2  'vbInteger 2 Integer
                .DataType = 5
                tmpInt = CInt(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpLong), CLng(2)
                VarOutput(i) = .bytTagData ' CInt(.Data)
            Case Is = 3  'vbLong 3 Long integer
                .DataType = 3
                tmpLong = CLng(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpLong), CLng(4)
                VarOutput(i) = .bytTagData ' CLng(.Data)
            Case Is = 4  'vbSingle 4 Single-precision floating-point number
                .DataType = 4
                tmpSingle = CSng(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpSingle), CLng(4)
                VarOutput(i) = .bytTagData ' CSng(.Data)
            Case Is = 5  'vbDouble 5 Double-precision floating-point number
                .DataType = 4
                tmpDouble = CDbl(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpDouble), CLng(8)
                VarOutput(i) = .bytTagData 'CDbl(.Data)
            Case Is = 6  'vbCurrency 6 Currency value
                .DataType = 4
                tmpDouble = CDbl(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpDouble), CLng(8)
                VarOutput(i) = .bytTagData 'CDbl(.Data)
            Case Is = 7  'vbDate 7 Date value
                .DataType = 4
                VarOutput(i) = .bytTagData 'CDate(.Data)
            Case Is = 8  'vbString 8 String
                .DataType = 0
                VarOutput(i) = .Data & Chr$(0)
                TempString(i) = StrConv(VarOutput(i), vbUnicode)
                .bytTagData = StrConv(TempString(i), vbFromUnicode)
            'Case Is = 9  'vbObject 9 Object
            'Case Is = 10  'vbError 10 Error value
            Case Is = 11  'vbBoolean 11 Boolean value
                'VB boolean is 2 bytes - Extended Content boolean is 4 bytes
                .DataType = 2
                tmpLong = CLng(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpLong), CLng(4)
                VarOutput(i) = .bytTagData
                Debug.Print LenB(VarOutput(i))
            'Case Is = 12  'vbVariant 12 Variant (used only witharrays of variants)
            'Case Is = 13  'vbDataObject 13 A data access object
            Case Is = 14  'vbDecimal 14 Decimal value
                .DataType = 5
                VarOutput(i) = .bytTagData ' CInt(.Data)
            Case Is = 17  'vbByte 17 Byte value
            Case Is = 36  'vbUserDefinedType 36 Variants that contain user-defined types
                .DataType = 6
                VarOutput(i) = .bytTagData ' CInt(.Data)
                MsgBox "Cannot handle vbUserDefinedType", vbCritical, "Error in BuildExtContent"
            Case Is = 8192, 8209 'vbArray 8192 Array
                .DataType = 1
                VarOutput(i) = .bytTagData
            Case Else
                .DataType = 1
                VarOutput(i) = .bytTagData
                MsgBox "Cannot handle " & VarType(.Data), vbCritical, "Error in BuildExtContent"
        '
        '0x0000  Unicode string. The data consists of a sequence of Unicode characters.
        '0x0001  BYTE array. The type of the data is implementation-specific.
        '0x0002  BOOL. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
        '              Only 0x0000 or 0x0001 are permitted values.
        '0x0003  DWORD. The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
        '0x0004  QWORD. The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
        '0x0005  WORD. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
        '0x0006  GUID. The data is 16 bytes long and should be interpreted as a 128-bit GUID.
        End Select
        '0 indicates a string
        dLen = LenB(VarOutput(i))
        If dLen = 0 Then
            .DataType = 0
            .DataLen = 2
        Else
            .DataLen = dLen ' LenB(VarOutput(i))
        End If
        ExtContSize = ExtContSize + LenB(.DataType)
        ExtContSize = ExtContSize + LenB(.DataLen)
        ExtContSize = ExtContSize + .DataLen
        Debug.Print "'"; """"; .Name; " = "; .Data; """"
        Debug.Print ".NameLen = "; .NameLen; " '("; Hex$(.NameLen); ")"
        Debug.Print ".Name = "; """"; .Name; """"
        Debug.Print "'VB Data Type = "; VarType(.Data); " => "; TypeName(.Data)
        Debug.Print ".DataType = "; .DataType
        Debug.Print ".DataLen = "; .DataLen; " '("; Hex$(.DataLen); ")"
        Debug.Print ".Data = "; .Data
        Debug.Print "ExtContSize ="; ExtContSize; " '("; Hex$(ExtContSize); ")"
        Debug.Print "-------------End of "; i; "--------------------"
    End With
Next i
For i = 0 To DescriptorCount - 1
    With Descriptor(i)
        Debug.Print "Ext. Content for"; i; " = "; .Name; " => "; .DataLen; " = "; .Data
    End With
Next i
ExtContSize = ExtContSize + 26
Debug.Print "ExtContSize   ="; ExtContSize; " ("; Hex$(ExtContSize); ")"
Debug.Print "ExtCont_QSize ="; ExtCont_QSize; " ("; Hex$(ExtCont_QSize); ")"
With objExtContInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Extended_Content_Description_Object)
    .ObjectSize.High = HiDWORD(ExtContSize)
    .ObjectSize.Low = LoDWORD(ExtContSize)
    .RecordCount = DescriptorCount
End With
If Len(eM$) > 0 Then
    Debug.Print eM$
    MsgBox eM$, , "ASF_File_Properties_Object"
    eM$ = ""
End If
'
GoTo Exit_BuildExtContent
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine BuildExtContent "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in BuildExtContent"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
Debug.Assert (mError <> vbRetry)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_BuildExtContent:
End Sub

Sub PutExtContentData()
On Error GoTo Oops
Dim dLen As Long
'
Debug.Print "---------------------------"
Debug.Print "Saving objExtContInfo..."
'
Put #fNum, , objExtContInfo
'
Debug.Print "------------------Putting Data-----------------"
For i = 0 To DescriptorCount - 1
    With Descriptor(i)
        Put #fNum, , .NameLen
        Put #fNum, , .bytTagName
        Debug.Print VarType(.DataType); "=>"; TypeName(.DataType)
        Put #fNum, , .DataType
        Put #fNum, , .DataLen
    '    ReDim .bytTagData(0)
    '    If Descriptor.DataLen > 0 Then
    '        ReDim .bytTagData(.DataLen - 1)
    '        CopyMemory .bytTagData(0), ByVal VarPtr(.Data), CLng(.DataLen)
    '    Else
    '        ReDim .bytTagData(0)
    '    End If
    '    Put #fNum, , Descriptor.bytTagData
        If LenB(VarOutput(i)) > 0 Then
            Put #fNum, , .bytTagData 'VarOutput(i)
        End If
    End With
Next i
HeaderCompleteCount = HeaderCompleteCount + 1
GoTo Exit_PutExtContent
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutExtContent "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutExtContent"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
Debug.Assert (mError <> vbRetry)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutExtContent:
End Sub

Sub BuildFileProperties()
Debug.Print "---------------------------"
Debug.Print "Building objFileInfo..."
With objFileInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_File_Properties_Object)
    .ObjectSize.High = HiDWORD(FP_QSize)
    .ObjectSize.Low = LoDWORD(FP_QSize)
    Debug.Print "ObjectSize             = "; FP_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    .FileID = GUIDStrToGUIDStruct(strFile_GUID)
    Debug.Print "File GUID = "; strFile_GUID
    '
    .FileSize.High = HiDWORD(calcFileQSize)
    .FileSize.Low = LoDWORD(calcFileQSize)
    Debug.Print "FileSize               = "; calcFileQSize; "("; Hex$(.FileSize.High); AddZero(Hex$(.FileSize.Low), 8); ")"
End With
End Sub

Sub PutFileProperties()
On Error GoTo Oops
'------------------------------ASF_File_Properties_Object----------------------------------------
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutFileProperties"
Debug.Print "---------------------------"
Debug.Print "Saving objFileInfo..."
'
Put #fNum, , objFileInfo
HeaderCompleteCount = HeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutFileProperties"
GoTo Exit_PutFileProperties
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutFileProperties "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutFileProperties"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutFileProperties:
End Sub

Sub BuildHeaderExtension()
Debug.Print "-------------------ASF_Header_Extension_Object-----------------------------"
eM$ = ""
Debug.Print "Building objHeaderExtInfo..."
ExtDataBytesToWrite = 0
ExtDataBytesToWrite = ExtDataBytesToWrite + Comp_QSize
ExtDataBytesToWrite = ExtDataBytesToWrite + MetaData_QSize
ExtDataBytesToWrite = ExtDataBytesToWrite + Pad_QSize
ExtDataBytesToWrite = ExtDataBytesToWrite + ExtStream_QSize
ExtDataBytesToWrite = ExtDataBytesToWrite + PH_QSize
ExtDataBytesToWrite = ExtDataBytesToWrite + MetaDataLib_QSize
With objHeaderExtInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Header_Extension_Object)
    Ext_QSize = Len(objHeaderExtInfo) + ExtDataBytesToWrite
    .ObjectSize.High = HiDWORD(Ext_QSize)
    .ObjectSize.Low = LoDWORD(Ext_QSize)
    Debug.Print "ObjectSize         = "; Ext_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    .Reserved1 = GUIDStrToGUIDStruct(ASF_Reserved_1)
    .Reserved2 = 6
    '
    .ExtDataSize = ExtDataBytesToWrite
    'Specifies the number of bytes stored in the Header Extension Data field.
    'This value may be 0 bytes or 24 bytes and larger.
    'It should also be equal to the Object Size field minus 46 bytes.
    If .ExtDataSize <> Ext_QSize - 46 Then
        eM$ = "There may be an error with the size of the extended data..."
        eM$ = eM$ & vbCrLf & "ObjectSize    = " & Ext_QSize & " (" & Hex$(Ext_QSize) & ")"
        eM$ = eM$ & vbCrLf & "ExtDataSize   = " & .ExtDataSize & " (" & Hex$(.ExtDataSize) & ")"
    Else
        Debug.Print "ExtDataSize OK     = "; .ExtDataSize; "("; Hex$(.ExtDataSize); ") = Ext_QSize - 46 "
    End If
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Header_Extension_Object"
    eM$ = ""
End If
End Sub

Sub PutHeaderExtension()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objHeaderExtInfo"
eM$ = ""
'
Debug.Print "Saving objHeaderExtInfo..."
'
With objHeaderExtInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Header_Extension_Object)
    Ext_QSize = Len(objHeaderExtInfo) + .ExtDataSize
    .ObjectSize.High = HiDWORD(Ext_QSize)
    .ObjectSize.Low = LoDWORD(Ext_QSize)
    Debug.Print "ObjectSize         = "; Ext_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    .Reserved1 = GUIDStrToGUIDStruct(ASF_Reserved_1)
    .Reserved2 = 6
    '
    'Specifies the number of bytes stored in the Header Extension Data field.
    'This value may be 0 bytes or 24 bytes and larger.
    'It should also be equal to the Object Size field minus 46 bytes.
    If .ExtDataSize <> Ext_QSize - 46 Then
        eM$ = "There may be an error with the size of the extended data..."
        eM$ = eM$ & vbCrLf & "ObjectSize    = " & Ext_QSize & " (" & Hex$(Ext_QSize) & ")"
        eM$ = eM$ & vbCrLf & "ExtDataSize   = " & .ExtDataSize & " (" & Hex$(.ExtDataSize) & ")"
    Else
        Debug.Print "ExtDataSize OK     = "; .ExtDataSize; "("; Hex$(.ExtDataSize); ") = Ext_QSize - 46 "
    End If
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Header_Extension_Object"
    eM$ = ""
End If
'
Put #fNum, , objHeaderExtInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutHeaderExtension"
GoTo Exit_PutHeaderExtension
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutHeaderExtension "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutHeaderExtension"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutHeaderExtension:
End Sub

Sub BuildCompatibility()
Debug.Print "-------------------ASF_Undocumented_Compatibility_GUID-----------------------------"
eM$ = ""
'now add the objCompatibilityInfo length
Comp_QSize = Len(objCompatibilityInfo)
With objCompatibilityInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Undocumented_Compatibility_GUID)
    .ObjectSize.High = HiDWORD(Comp_QSize)
    .ObjectSize.Low = LoDWORD(Comp_QSize)
    Debug.Print "ObjectSize         = "; Comp_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    ExtDataBytesLeft = ExtDataBytesLeft - Comp_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    'The value of this field shall be set to 26 bytes
    If Comp_QSize <> 26 Then
        eM$ = eM$ & vbCrLf & "** Error in ObjectSize = " & Comp_QSize & " (" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ") - always 26"
    Else
        Debug.Print "ObjectSize OK  = "; Comp_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ") - always 26"
    End If
     .Profile = 2
    .Mode = 1
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "Compatibility Object"
    eM$ = ""
End If
End Sub

Sub PutCompatibility()
'Undocumented object...
Debug.Print "Saving objHeaderExtInfo..."
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutCompatibility"
Put #fNum, , objCompatibilityInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutCompatibility"
End Sub

Sub BuildMetaData()
Debug.Print "-------------------ASF_Metadata_Object-----------------------------"
eM$ = ""
'calculate the size of the metadata
MetaData_QSize = Len(objMetaDataInfo)
For i = 0 To objMetaDataInfo.RecordCount - 1
    MetaData_QSize = MetaData_QSize + Len(objMetaDataRecord) 'add 12 bytes
    'Put Length of the name data
    MetaData_QSize = MetaData_QSize + MetaDescriptor(i).NameLen
    MetaData_QSize = MetaData_QSize + MetaDescriptor(i).DataLen
Next i
With objMetaDataInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Metadata_Object)
    .ObjectSize.High = HiDWORD(MetaData_QSize)
    .ObjectSize.Low = LoDWORD(MetaData_QSize)
    Debug.Print "ObjectSize         = "; MetaData_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    ExtDataBytesLeft = ExtDataBytesLeft - MetaData_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "RecordCount= "; .RecordCount; "("; Hex$(.RecordCount); ")"
End With
End Sub

Sub PutMetaData()
Debug.Print "Saving objMetaDataInfo..."
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutMetaData"
'
Put #fNum, , objMetaDataInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutMetaData"
'
PutMetaDataRecords 0, objMetaDataInfo.RecordCount - 1
End Sub

Sub PutMetaDataRecords(StartRec As Integer, EndRec As Integer)
On Error GoTo Oops
Dim strInfo$
eM$ = ""
Debug.Print "Saving data for "; MaxMTags; " MetaData Entries."
'
Debug.Print "----------------------------------------------------------"
Dim tmpLong As Long
For i = StartRec To EndRec
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print LastAddress$; " Before PutMetaDataRecords"
    MTagIndex = i
    With objMetaDataRecord
        .Reserved = 0
        .StreamNumber = MetaDescriptor(MTagIndex).StreamNumber
        Debug.Print "StreamNumber    = "; .StreamNumber; "("; Hex$(.StreamNumber); ")"
        'A value of 0 in this field indicates that it applies to the whole file;
        ' otherwise, the entry applies only to the indicated stream number
        ' and must be between 1 and 127.
        '
        'Put Length of the name data
        .NameLen = MetaDescriptor(MTagIndex).NameLen
        .DataType = MetaDescriptor(MTagIndex).DataType
        .DataLen = MetaDescriptor(MTagIndex).DataLen
    End With
    Put #fNum, , objMetaDataRecord
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print LastAddress$; " After objMetaDataRecord " & i
    '
    With MetaDescriptor(MTagIndex)
        Debug.Print "MetaData NameLen = "; .NameLen
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before .bytTagName"
        'Put the name of the MTag
        TempString(0) = StrConv(.Name & Chr$(0), vbUnicode)
        .NameLen = Len(.Name) + 1 'Len(TempString(0))
        .bytTagName = StrConv(TempString(0), vbFromUnicode)
        Put #fNum, , .bytTagName
        '
        'Put the data type
        Debug.Print "MetaData DataType = "; .DataType; " ("; Hex$(.DataType); ")"
        'Put the data length
        Debug.Print "MetaData DataLen = "; .DataLen; " ("; Hex$(.DataLen); ")"
        '
        'Metadata Library data type
'        0x0000  Unicode string. The data consists of a sequence of Unicode characters.
'        0x0001  BYTE array. The type of the data is implementation-specific.
'        0x0002  BOOL. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'                      Only 0x0000 or 0x0001 are permitted values.
'        0x0003  DWORD. The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'        0x0004  QWORD. The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'        0x0005  WORD. The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'        0x0006  GUID. The data is 16 bytes long and should be interpreted as a 128-bit GUID.
       'the Tag data must be set according to its datatype
        ReDim .bytTagData(0 To .DataLen - 1)
        Dim tmpBool As Boolean
        Select Case .DataType
            Case Is = 0 'string values
                TempString(0) = StrConv(.Data & Chr$(0), vbUnicode)
                .bytTagData = StrConv(TempString(0), vbFromUnicode)
            Case Is = 1  'BYTE array
                .bytTagData = .Data
            Case 2 'Boolean  VB boolean is 2 bytes - MetaData is 4 bytes
                tmpBool = CBool(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpBool), CLng(.DataLen)
            Case 3 'Long
                tmpLong = CLng(.Data)
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(tmpLong), CLng(.DataLen)
            Case 4 'Date or Double...probably date, but i can't figure out how to convert it
                'Date variables are stored as IEEE 64-bit (8-byte) floating-point numbers that
                'represent dates ranging from 1 January 100 to 31 December 9999
                'and times from 0:00:00 to 23:59:59.
                'Any recognizable literal date values can be assigned to Date variables.
                'When othernumeric types are converted to Date, values to the left of the decimal
                'represent date information while values to the right of the decimal represent time.
                'Midnight is 0 and midday is 0.5.
                'Negative whole numbers represent dates before 30 December 1899.
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(.Data), CLng(.DataLen)
            Case 5 'integer
                CopyMemory ByVal VarPtr(.bytTagData(0)), ByVal VarPtr(.Data), CLng(.DataLen)
            Case 6 ' GUID. The data is 16 bytes long and should be interpreted as a 128-bit GUID.
                'MsgBox "Put a Guid in PutMetaDataRecords"
             Case Else
                MsgBox "Error in Datatype for PutMetaDataRecords"
        End Select
ShowMe:
        'Show the tag name and data
        strInfo$ = .Name & " = " & .Data
        Debug.Print strInfo$
        'eM$ = eM$ & strInfo$ & vbCrLf
        Select Case .Name
            Case "WM/TrackNumber"
                .Data = InfoTag.TrackNumber
        End Select
        'Put the MetaData Data
        If .DataLen > 0 Then
            If .DataType = 6 Then 'guid
                GUIDOut = GUIDStrToGUIDStruct(CStr(.Data))
                Put #fNum, , GUIDOut
            Else
                Put #fNum, , .bytTagData
            End If
            'Debug.Print MetaDescriptor(MTagIndex).bytTagData
        End If
    End With
Next i
'
Debug.Print "----------------------------------------------------------"
'
GoTo Exit_PutMetaDataRecords
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutMetaDataRecords "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutMetaDataRecords"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutMetaDataRecords:
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutMetaDataRecords"
End Sub

Sub BuildPadding()
eM$ = ""
Debug.Print "-------------------ASF_Padding_Object-----------------------------"
'calculate the padding length be taking the original file's padding end point and
'subtracting our current position.  This will take into account changes we have made.
'PaddingLen = (PaddingEndPos - Seek(fNum)) - 24
PaddingLen = ((PaddingEndPos - 1) - AccumWrite_QSize) - Len(objPaddingInfo)
Pad_QSize = Len(objPaddingInfo) + PaddingLen
With objPaddingInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Padding_Object)
    .ObjectSize.High = HiDWORD(Pad_QSize)
    .ObjectSize.Low = LoDWORD(Pad_QSize)
    ExtDataBytesLeft = ExtDataBytesLeft - Pad_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize = "; Pad_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
End With
End Sub

Sub PutPadding()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutPadding"
'should be 1806D474-CADF-4509-A4BA-9AABCB96AAE8
Put #fNum, , objPaddingInfo
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objPaddingInfo"
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
ReDim bytPaddingData(PaddingLen - 1)
Put #fNum, , bytPaddingData
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After bytPaddingData"
GoTo Exit_PutPadding
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutPadding "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutPadding"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutPadding:
End Sub

Sub BuildExtendedStream()
Debug.Print "-------------------ASF_Extended_Stream_Properties_Object-----------------------------"
'the next guid should be ASF_Extended_Stream_Properties_Object
eM$ = ""
Dim varAvgTime As Variant
'add the extended stream object
ExtStream_QSize = Len(objExtStreamInfo)
With objExtStreamInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Extended_Stream_Properties_Object)
    .ObjectSize.High = HiDWORD(ExtStream_QSize)
    .ObjectSize.Low = LoDWORD(ExtStream_QSize)
    ExtDataBytesLeft = ExtDataBytesLeft - ExtStream_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize             = "; ExtStream_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    .StartTime.High = HiDWORD(StartTimeQSize)
    .StartTime.Low = LoDWORD(StartTimeQSize)
    Debug.Print "StartTime              = "; StartTimeQSize; "("; Hex$(.StartTime.High); AddZero(Hex$(.StartTime.Low), 8); ")"
    '
    .EndTime.High = HiDWORD(EndTimeQSize)
    .EndTime.Low = LoDWORD(EndTimeQSize)
    Debug.Print "EndTime                = "; EndTimeQSize; "("; Hex$(.EndTime.High); AddZero(Hex$(.EndTime.Low), 8); ")"
    '
    .AvgTimePerFrame.High = HiDWORD(AvgTimePerFrameQSize)
    .AvgTimePerFrame.Low = LoDWORD(AvgTimePerFrameQSize)
    Debug.Print "AvgTimePerFrame        = "; AvgTimePerFrameQSize; "("; Hex$(.AvgTimePerFrame.High); AddZero(Hex$(.AvgTimePerFrame.Low), 8); ")"
    varAvgTime = CVar(AvgTimePerFrameQSize * ((10 ^ -7)))
    Debug.Print "AvgTimePerFrame (sec.) = "; varAvgTime
    '
    Debug.Print "StreamNameCount        = "; .StreamNameCount; "("; Hex$(.StreamNameCount); ")"
    Debug.Print "PayloadExtSystemCount  = "; .PayloadExtSystemCount; "("; Hex$(.PayloadExtSystemCount); ")"
End With
End Sub

Sub PutExtendedStream()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutExtendedStream"
'should be 1806D474-CADF-4509-A4BA-9AABCB96AAE8
Put #fNum, , objExtStreamInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objExtStreamInfo"
'Stream Names
'The structure of each Stream Name entry is shown in the following table.
'Field name             Field type  Size (bytes)
'-----------------------------------------------------------------------------
'Language ID Index      WORD        2
'Stream Name Length     WORD        2
'Stream Name            WCHAR   varies
If objExtStreamInfo.StreamNameCount > 0 Then
    For i = 0 To objExtStreamInfo.StreamNameCount - 1
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before Put ExtStreamName"
        Put #fNum, , LangIDIndex
        StreamNLen = Len(StreamName(i))
        Put #fNum, , StreamNLen
        ReDim bytStreamName(StreamNLen)
        Put #fNum, , bytStreamName
        Debug.Print i; ") StreamName = "; StreamName(i)
    Next i
End If
'
'Payload extension systems are described in the following table.
'Field name  Field type  Size (bits)
''-----------------------------------------------------------------------------
'Extension System ID             GUID    16
'Extension Data Size             WORD    2
'Extension System Info Length    DWORD   4
'Extension System Info           BYTE    varies
'below is an array for each ExtSystemInfo
'ASF_Payload_Extension_System_Timecode'             399595EC-8667-4E2D-8FDB-98814CE76C1E
'ASF_Payload_Extension_System_File_Name             E165EC0E-19ED-45D7-B4A7-25CBD1E28E9B
'ASF_Payload_Extension_System_Content_Type          D590DC20-07BC-436C-9CF7-F3BBFBF1A4DC
'ASF_Payload_Extension_System_Pixel_Aspect_Ratio    1B1EE554-F9EA-4BC8-821A-376B74E4C4B8
'ASF_Payload_Extension_System_Sample_Duration       C6BD9450-867F-4907-83A3-C77921B733AD
'ASF_Payload_Extension_System_Encryption_Sample_ID  6698B84E-0AFA-4330-AEB2-1C0A98D7A44D
If objExtStreamInfo.PayloadExtSystemCount > 0 Then
    For i = 0 To objExtStreamInfo.PayloadExtSystemCount - 1
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before objPayloadInfo"
        With objPayloadInfo
            .ExtSystemID = GUIDStrToGUIDStruct(strExtSystem_GUID)
            ExtStream_QSize = Len(objPaddingInfo)
            Debug.Print "objPayloadInfo GUID = "; strExtSystem_GUID
            Debug.Print "ExtDataSize        = "; .ExtDataSize; "("; Hex$(.ExtDataSize); ")"
            Debug.Print "ExtSystemInfoLen   = "; .ExtSystemInfoLen; "("; Hex$(.ExtSystemInfoLen); ")"
            'below puts the bytes for each ExtSystemInfo
        End With
        Put #fNum, , objPayloadInfo
        Put #fNum, , bytExtSystemInfo
    Next i
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After objPayloadInfo"
Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
GoTo Exit_PutExtendedStream
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutExtendedStream "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutExtendedStream"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutExtendedStream:
End Sub


Sub BuildPlaceholder()
'should be ASF_Index_Placeholder_Object = "D9AADE20-7C17-4F9C-BC28-8555DD98E2A2"
Debug.Print "-------------------ASF_Index_Placeholder_Object-----------------------------"
PH_QSize = Len(objPlaceholderInfo)
With objPlaceholderInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Index_Placeholder_Object)
    .ObjectSize.High = HiDWORD(PH_QSize)
    .ObjectSize.Low = LoDWORD(PH_QSize)
    ExtDataBytesLeft = ExtDataBytesLeft - PH_QSize
    Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
    Debug.Print "ObjectSize = "; PH_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    .Word1 = 0  'always 0
    .Word2 = 0  'always 0
    .Word3 = 0  'always 0
    .Word4 = 0  'always 0
    .Word5 = 0 'always 0
End With
End Sub

Sub PutPlaceholder()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutPlaceholder"
Put #fNum, , objPlaceholderInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutPlaceholder"
Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
GoTo Exit_PutPlaceholder
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutPlaceholder "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutPlaceholder"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutPlaceholder:
End Sub

Sub BuildMetaDataLibrary()
On Error GoTo Oops
'calculate the size of the metadata for the MetaDataLibrary
Dim RecIndex As Integer
'on the temp files, the metadata library is empty
'on the stored files I have it is:
'MetaData Library Entries
'4=>WM/MediaClassPrimaryID
'Reserved     =0
'StreamNumber =0
'NameLen      =46
'Name         =WM/MediaClassPrimaryID
'DataLen      =16
'DataType     =6
'Data         =D1607DBC-E323-4BE2-86A1-48A42A28441E
'
'5=>WM/MediaClassSecondaryID
'Reserved     =0
'StreamNumber =0
'NameLen      =50
'Name         =WM/MediaClassSecondaryID
'DataLen      =16
'DataType     =6
'Data         =00000000-0000-0000-0000-000000000000
'
'6=>WM/WMContentID
'Reserved     =0
'StreamNumber =0
'NameLen      =30
'Name         =WM/WMContentID
'DataLen      =16
'DataType     =6
'Data         =16DA729D-4284-43D9-AEC4-B5F44B87B227
'
'7=>WM/WMCollectionID
'Reserved     =0
'StreamNumber =0
'NameLen      =36
'Name         =WM/WMCollectionID
'DataLen      =16
'DataType     =6
'Data         =E982F8DB-C749-449C-984F-A5C7B02499C3
'
'8=>WM/WMCollectionGroupID
'Reserved     =0
'StreamNumber =0
'NameLen      =46
'Name         =WM/WMCollectionGroupID
'DataLen      =16
'DataType     =6
'Data         =E982F8DB-C749-449C-984F-A5C7B02499C3
'
MetaDataLib_QSize = Len(objMetaDataLibraryInfo)
If objMetaDataLibraryInfo.RecordCount = 0 Then
    objMetaDataLibraryInfo.RecordCount = 5
    ReDim Preserve MetaDescriptor(MaxMTags + 5)
    MTagIndex = MaxMTags
    With MetaDescriptor(MTagIndex)
        .Reserved = 0
        .StreamNumber = 0
        .Name = "WM/MediaClassPrimaryID"
        .Data = "D1607DBC-E323-4BE2-86A1-48A42A28441E"
        .NameLen = 2 * (Len(.Name & Chr$(0)))
        .DataType = 6
        .DataLen = 16
    End With
    MTagIndex = MTagIndex + 1
    With MetaDescriptor(MTagIndex)
        .Reserved = 0
        .StreamNumber = 0
        .Name = "WM/MediaClassSecondaryID"
        .Data = "00000000-0000-0000-0000-000000000000"
        .NameLen = 2 * (Len(.Name & Chr$(0)))
        .DataType = 6
        .DataLen = 16
    End With
    MTagIndex = MTagIndex + 1
    With MetaDescriptor(MTagIndex)
        .Reserved = 0
        .StreamNumber = 0
        .Name = "WM/WMContentID"
        .Data = "16DA729D-4284-43D9-AEC4-B5F44B87B227"
        .NameLen = 2 * (Len(.Name & Chr$(0)))
        .DataType = 6
        .DataLen = 16
    End With
    MTagIndex = MTagIndex + 1
    With MetaDescriptor(MTagIndex)
        .Reserved = 0
        .StreamNumber = 0
        .Name = "WM/WMCollectionID"
        .Data = "E982F8DB-C749-449C-984F-A5C7B02499C3"
        .NameLen = 2 * (Len(.Name & Chr$(0)))
        .DataType = 6
        .DataLen = 16
    End With
    MTagIndex = MTagIndex + 1
    With MetaDescriptor(MTagIndex)
        .Reserved = 0
        .StreamNumber = 0
        .Reserved = 0
        .StreamNumber = 0
        .Name = "WM/WMCollectionGroupID"
        .Data = "E982F8DB-C749-449C-984F-A5C7B02499C3"
        .NameLen = 2 * (Len(.Name & Chr$(0)))
        .DataType = 6
        .DataLen = 16
    End With
    MaxMTags = MaxMTags + objMetaDataLibraryInfo.RecordCount
End If
For i = 0 To objMetaDataLibraryInfo.RecordCount - 1
    MetaDataLib_QSize = MetaDataLib_QSize + Len(objMetaDataRecord) 'add 12 bytes
    'calculate Length of the name data
    RecIndex = i + objMetaDataInfo.RecordCount
    MetaDataLib_QSize = MetaDataLib_QSize + MetaDescriptor(RecIndex).NameLen
    MetaDataLib_QSize = MetaDataLib_QSize + MetaDescriptor(RecIndex).DataLen
Next i
With objMetaDataLibraryInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Metadata_Library_Object)
    .ObjectSize.High = HiDWORD(MetaDataLib_QSize)
    .ObjectSize.Low = LoDWORD(MetaDataLib_QSize)
    ExtDataBytesLeft = ExtDataBytesLeft - MetaDataLib_QSize
    Debug.Print "ObjectSize = "; MetaDataLib_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    Debug.Print "RecordCount= "; .RecordCount; "("; Hex$(.RecordCount); ")"
End With
GoTo Exit_BuildMetaDataLibrary
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine BuildMetaDataLibrary "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in BuildMetaDataLibrary"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
'Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_BuildMetaDataLibrary:
End Sub

Sub PutMetadataLibrary()
On Error GoTo Oops
'we also use this structure for ASF_Metadata_Object
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutMetadataLibrary"
Debug.Print "-------------------ASF_Metadata_Library_Object-----------------------------"
eM$ = ""
Put #fNum, , objMetaDataLibraryInfo
SubHeaderCompleteCount = SubHeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutMetadataLibrary"
Debug.Print "ExtDataBytesLeft = "; ExtDataBytesLeft
'
PutMetaDataRecords objMetaDataInfo.RecordCount, objMetaDataInfo.RecordCount + objMetaDataLibraryInfo.RecordCount - 1
GoTo Exit_PutMetadataLibrary
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutMetadataLibrary "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutMetadataLibrary"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutMetadataLibrary:
End Sub

Sub BuildCodecs()
Codec_QSize = Len(objCodecListInfo)
'calculate the size of the Codec data
For i = 0 To CodecCount - 1
    Codec_QSize = Codec_QSize + 8 ' Len(Codec(i))
    Codec_QSize = Codec_QSize + (2 * Codec(i).NameLen)
    Codec_QSize = Codec_QSize + (2 * Codec(i).DescrLen)
    Codec_QSize = Codec_QSize + Codec(i).DataLen
Next i
End Sub

Sub PutCodecList()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutCodecList"
eM$ = ""
Debug.Print "-------------------ASF_Codec_List_Object-----------------------------"
With objCodecListInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Codec_List_Object)
    .ObjectSize.High = HiDWORD(Codec_QSize)
    .ObjectSize.Low = LoDWORD(Codec_QSize)
    Debug.Print "ObjectSize         = "; Codec_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    .Reserved = GUIDStrToGUIDStruct(ASF_Reserved_2)
    'should be = ASF_Reserved_2 - 86D15241-311D-11D0-A3A4-00A0C90348F6
    Debug.Print "CodecEntriesCount  = "; .CodecEntriesCount; "("; Hex$(.CodecEntriesCount); ")"
End With
If Len(eM$) > 0 Then
    MsgBox eM$, , "ASF_Codec_List_Object"
    eM$ = ""
End If
Put #fNum, , objCodecListInfo
HeaderCompleteCount = HeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutCodecList"
PutCodecs
GoTo Exit_PutCodecList
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutCodecList "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutCodecList"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutCodecList:
End Sub

Sub PutCodecs()
On Error GoTo Oops
ReDim TempString(2)
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutCodecs"
CodecCount = objCodecListInfo.CodecEntriesCount
Debug.Print "Writing data for "; CodecCount; " Codec Entries."
eM$ = ""
'
'Field Name             Field Type  Size (bytes)
'----------------------------------------------
'Type                       WORD    2
'Codec Name Length          WORD    2
'Codec Name                 WCHAR  varies
'Codec Description Length   WORD    2
'Codec Description          WCHAR  varies
'Codec Information Length   WORD    2
'Codec Information          BYTE   varies
'
For i = 0 To CodecCount - 1
    CTagIndex = i 'MaxCTags + i
    With Codec(CTagIndex)
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before Codec Data " & i
        Put #fNum, , .Type
        If .Type <> 2 Then
            MsgBox "This is not an audio codec!"
        End If
        Debug.Print "Codec Type = "; .Type
        'Put Length of the name data
        LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
        Debug.Print LastAddress$; " Before Codec Name " & i
        TempString(0) = StrConv(.Name & Chr$(0), vbUnicode)
        .NameLen = Len(.Name) + 1 'Len(TempString(0))
        .bytCodecName = StrConv(TempString(0), vbFromUnicode)
        Debug.Print .Name
        Put #fNum, , .NameLen
        Debug.Print "Codec NameLen = "; .NameLen
        Put #fNum, , .bytCodecName
        '
        'Put the Description length
        If Len(.Descr) > 0 Then
            TempString(1) = StrConv(.Descr & Chr$(0), vbUnicode)
            .DescrLen = Len(.Descr) + 1 'Len(TempString(1))
            .bytCodecDescr = StrConv(TempString(1), vbFromUnicode)
        Else
            .DescrLen = 0
        End If
        Debug.Print .Descr
        Put #fNum, , .DescrLen
        If .DescrLen > 0 Then
            Put #fNum, , .bytCodecDescr
        End If
        '
        'Put the Codec Data
        .DataLen = LenB(.Data)
        Put #fNum, , .DataLen
        If .DataLen > 0 Then
            Put #fNum, , .Data
            Debug.Print "(Format is "; GetFormatTag(CInt(.Data)); ")"
        End If
    End With
Next i
'
GoTo Exit_PutCodecs
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutCodecs "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutCodecs"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutCodecs:
End Sub

Sub BuildStreamProperties()
With objStreamInfo
    Debug.Print "-------------------ASF_Stream_Properties_Object-------------------------"
    .ObjectID = GUIDStrToGUIDStruct(ASF_Stream_Properties_Object)
    'ObjectSize is calculated below....
    strStreamType_GUID = ASF_Audio_Media
    .StreamType = GUIDStrToGUIDStruct(strStreamType_GUID) 'ASF_Audio_Media
    .ErrorType = GUIDStrToGUIDStruct(ASF_Audio_Spread)
    '
    .TimeOffset.High = HiDWORD(TimeOffsetQSize)
    .TimeOffset.Low = LoDWORD(TimeOffsetQSize)
    Debug.Print "TimeOffset     = "; TimeOffsetQSize; "("; Hex$(.TimeOffset.High); AddZero(Hex$(.TimeOffset.Low), 8); ")"
    Dim varTimeOffset As Variant
    varTimeOffset = CVar(TimeOffsetQSize * ((10 ^ -7)))
    Debug.Print "TimeOffset     = "; varTimeOffset; " sec."
    If strStreamType_GUID = ASF_Audio_Media Then
        .TypeDataLen = Len(objAudioInfo) + objAudioInfo.CodecDataSize
    Else
        .TypeDataLen = Len(TypeSpecificData)
    End If
    Debug.Print "TypeDataLen    = "; .TypeDataLen; "("; Hex$(.TypeDataLen); ")"
    .ErrorDataLen = LenB(ErrorCorrectionData)
    Debug.Print "ErrorDataLen   = "; .ErrorDataLen; "("; Hex$(.ErrorDataLen); ")"
    Debug.Print "Flags          = "; .Flags; "("; Hex$(.Flags); ")"
    .Reserved = .Reserved 'supposed to be 0 odd...it seems to be NEVER be 0
    'calculate the total size
    Stream_QSize = Len(objStreamInfo) + .TypeDataLen + .ErrorDataLen 'len of stream
    .ObjectSize.High = HiDWORD(Stream_QSize)
    .ObjectSize.Low = LoDWORD(Stream_QSize)
    Debug.Print "ObjectSize       = "; Stream_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
End With
End Sub

Sub PutStreamProperties()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutStreamProperties"
Put #fNum, , objStreamInfo
HeaderCompleteCount = HeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutStreamProperties"
'
'we need to jump to the audio parser
If strStreamType_GUID = ASF_Audio_Media Then
    PutAudio
Else
    Put #fNum, , TypeSpecificData
End If
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After Put TypeSpecificData"
'Error Correction Data
Put #fNum, , ErrorCorrectionData
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After Put ErrorCorrectionData"
GoTo Exit_PutStreamProperties
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutStreamProperties "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutStreamProperties"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutStreamProperties:
End Sub

Sub PutAudio()
On Error GoTo Oops
ASF_Audio_Media:
'ASF_Stream_Bitrate_Properties_Object
'then we dig out the ASF_Audio_Media
'F8699E40-5B4D-11CF-A8FD-00805F5C442B
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutAudio"
'
Put #fNum, , objAudioInfo 'ASF_Audio_Media
With objAudioInfo
    Debug.Print "-------------------ASF_Audio_Media-----------------------------"
    Debug.Print "CodecID        = "; .CodecID; "("; Hex$(.CodecID); ")"
    Debug.Print "(Format is "; GetFormatTag(.CodecID); ")"
    Debug.Print "NumChannels    = "; .NumChannels; "("; Hex$(.NumChannels); ")"
    Debug.Print "SamplesPerSec  = "; .SamplesPerSec; "("; Hex$(.SamplesPerSec); ")"
    Debug.Print "AvgBytesPerSec = "; .AvgBytesPerSec; "("; Hex$(.AvgBytesPerSec); ")"
    Debug.Print "BlockAlignment = "; .BlockAlignment; "("; Hex$(.BlockAlignment); ")"
    Debug.Print "BitsPerSample  = "; .BitsPerSample; "("; Hex$(.BitsPerSample); ")"
    Debug.Print "CodecDataSize  = "; .CodecDataSize; "("; Hex$(.CodecDataSize); ")"
    Debug.Print "CodecData"
    If .CodecDataSize > 0 Then
        Put #fNum, , bytCodecSpecificData
    End If
End With
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutAudio"
GoTo Exit_PutAudio
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutAudio "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutAudio"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutAudio:
End Sub

Sub BuildStreamBitrate()
With objBitRateInfo
    BitRate_QSize = Len(objBitRateInfo) + (.BitrateRecords * Len(BitRateEntry(0)))
End With
End Sub

Sub PutStreamBitrate()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before PutStreamBitrate"
With objBitRateInfo
    .ObjectID = GUIDStrToGUIDStruct(ASF_Stream_Bitrate_Properties_Object)
    'calculate the total size
    .ObjectSize.High = HiDWORD(BitRate_QSize)
    .ObjectSize.Low = LoDWORD(BitRate_QSize)
    Debug.Print "ObjectSize       = "; BitRate_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    '
    Debug.Print "BitrateRecords    = "; .BitrateRecords; "("; Hex$(.BitrateRecords); ")"
End With
Put #fNum, , objBitRateInfo
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After PutStreamBitrate"
'
For i = 0 To objBitRateInfo.BitrateRecords - 1
    LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
    Debug.Print LastAddress$; " Before BitrateRecord"
    With BitRateEntry(i)
        Debug.Print "Bitrate Flags = "; .Flags; "("; Hex$(.Flags); ")"
        Debug.Print "AvgBitrate    = "; .AvgBitrate; "("; Hex$(.AvgBitrate); ")"
    End With
    Put #fNum, , BitRateEntry(i)
Next i
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After All BitrateRecords"
GoTo Exit_PutStreamBitrate
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutStreamBitrate "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutStreamBitrate"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutStreamBitrate:
End Sub

Sub PutData()
On Error GoTo Oops
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " Before objDataInfo"
With objDataInfo
    Debug.Print "----------------------ASF_Data_Object-----------------------------"
    '
    'objDataInfo
    '    ObjectID As GuidStruct 'should be 75B22636-668E-11CF-A6D9-00AA0062CE6C
    '    ObjectSize As DlngSize 'at least 50 bytes
    '    FileID As GuidStruct
    '    TotalDataPackets As DlngSize 'must be equal to the Data Packet Count field in the File Properties Object
    '    Reserved As Integer
    'then the data packets
    'calculate the total size
    'Data_QSize = Len(objDataInfo) +
    AccumWrite_QSize = AccumWrite_QSize + Data_QSize
    frmMain.txtInfo.SelText = AccumWrite_QSize & vbCrLf
    .ObjectID = GUIDStrToGUIDStruct(ASF_Data_Object)
    .ObjectSize.High = HiDWORD(Data_QSize)
    .ObjectSize.Low = LoDWORD(Data_QSize)
    Debug.Print "ObjectSize       = "; Data_QSize; "("; Hex$(.ObjectSize.High); AddZero(Hex$(.ObjectSize.Low), 8); ")"
    If strDataFileID_GUID <> strFile_GUID Then
        eM$ = eM$ & vbCrLf & "*** Error in FileID *** "
        eM$ = eM$ & vbCrLf & "It must be identical to the value of the File ID field of the Header Object."
        eM$ = eM$ & vbCrLf & "strDataFileID_GUID  = " & strDataFileID_GUID
        eM$ = eM$ & vbCrLf & "FP strFile_GUID = " & strFile_GUID
        MsgBox eM$, , "ASF_Header_Object"
        eM$ = ""
    Else
        Debug.Print "strDataFileID_GUID  = "; strDataFileID_GUID
    End If
    .FileID = GUIDStrToGUIDStruct(strDataFileID_GUID)
    Debug.Print "FileID GUID = "; strDataFileID_GUID
    '
    Packets_QSize = Packets_QSize
    .TotalDataPackets.High = HiDWORD(Packets_QSize)
    .TotalDataPackets.Low = LoDWORD(Packets_QSize)
    Debug.Print "TotalDataPackets  = "; Packets_QSize; "("; Hex$(.TotalDataPackets.High); AddZero(Hex$(.TotalDataPackets.Low), 8); ")"
    .Reserved = 257
End With
Put #fNum, , objDataInfo 'ASF_Data_Object
HeaderCompleteCount = HeaderCompleteCount + 1
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After Put objDataInfo"
'
HeaderEnd = Seek(fNum)
Debug.Print "HeaderEnd = "; HeaderEnd
Put #fNum, HeaderEnd, StreamData  'this puts the whole rest of the file
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After Stream Data"
'
EndOfDataAddress = Seek(fNum) - 1
If (calcFileQSize = AccumRead_QSize) And (AccumRead_QSize = LOF(fNum)) And (LOF(fNum) = (EndOfDataAddress)) Then
    'everything matches...yay
    Debug.Print "All sizes match..."
Else
    eM$ = eM$ & vbCrLf & "*** Error in FileSize ***"
    eM$ = eM$ & vbCrLf & "Actual FileSize        = " & InfoTag.FileSize & vbTab & " (" & Hex$(InfoTag.FileSize) & ")"
    eM$ = eM$ & vbCrLf & "FP Object FileSize     = " & calcFileQSize & vbTab & " (" & Hex$(calcFileQSize) & ")"
    eM$ = eM$ & vbCrLf & "Accum. Read FileSize   = " & AccumRead_QSize & vbTab & " (" & Hex$(AccumRead_QSize) & ")"
    eM$ = eM$ & vbCrLf & "Accum. Write FileSize  = " & AccumWrite_QSize & vbTab & " (" & Hex$(AccumWrite_QSize) & ")"
    eM$ = eM$ & vbCrLf & "Last Read address  = " & EndOfDataAddress & vbTab & " (" & Hex$(EndOfDataAddress) & ")"
End If
'
Debug.Print "-------------Below 5 should be the same---------------"
Debug.Print LOF(fNum), AddZero(Hex$(LOF(fNum)), 6); "H is the Length FileSize"
Debug.Print calcFileQSize, AddZero(Hex$(calcFileQSize), 6); "H is the calcFileQSize from File Properties"
Debug.Print AccumRead_QSize, AddZero(Hex$(AccumRead_QSize), 6); "H is the AccumRead_QSize"
Debug.Print AccumWrite_QSize, AddZero(Hex$(AccumWrite_QSize), 6); "H is the AccumWrite_QSize"
Debug.Print EndOfDataAddress, AddZero(Hex$((EndOfDataAddress)), 6); "H is LastAddress"
'
Debug.Print ASF_QSize, AddZero(Hex$(ASF_QSize), 6); "H is the Header length"
Debug.Print "-------------Below should be the same---------------"
Debug.Print LOF(fNum) - ASF_QSize, AddZero(Hex$(LOF(fNum) - ASF_QSize), 6); " is the Size - Header length"
Debug.Print Data_QSize, AddZero(Hex$(Data_QSize), 6); " is the size of the data"
'
GoTo Exit_PutData
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine PutData "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutData"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutData:
End Sub

Sub TestNextGUID()
StartAddr = Seek(fNum)
LastAddress$ = AddZero(Hex$(StartAddr), 6) & "H"
Debug.Print LastAddress$; " Before TestNextGUID"
Get #fNum, , NextGUID
GUIDIn = NextGUID
strNextGUID = GUIDToString(GUIDIn)
Debug.Print "Next GUID = ";
Debug.Print ReturnGUIDName(strNextGUID)
Debug.Print strNextGUID
Debug.Print "------------------------------------------------"
Seek fNum, StartAddr
End Sub

Sub GetGUID()
Get #fNum, , NextGUID
GUIDIn = NextGUID
strGUID = GUIDToString(GUIDIn)
Debug.Print "GUID = "; strGUID
LastAddress$ = AddZero(Hex$(Seek(fNum)), 6) & "H"
Debug.Print LastAddress$; " After GetGUID"
End Sub

