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
'from a goup
'ASF files are logically composed of three types of top-level objects:
'the Header Object, the Data Object, and the Index Object(s).
'The Header Object is mandatory and must be placed at the beginning of every ASF file.
'The Data Object is also mandatory and must follow the Header Object.
'The Index Object(s) are optional, but they are useful in providing time-based
'random access into ASF files. When present, the Index Object(s) must be the
'last object in the ASF file.
'
'=====================================================================================
'The Header Object is represented using the following structure.
'-----------------------------------------------
'Field name              Field type   Size (bytes)
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
'The Content Description Object is represented using the following structure.
'
'Field name             Field type  Size (bytes)
'Object ID              GUID        16
'Object Size            QWORD       8
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
'0x0002  BOOL.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer. Only 0x0000 or 0x0001 are permitted values.
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
fixthis
'The Stream Properties Object is represented using the following structure.
'
'Field Name Field Type  Size (bits)
'Object ID  GUID    128
'Object Size    QWORD   64
'Stream Type    GUID    128
'Error Correction Type  GUID    128
'Time Offset    QWORD   64
'Type-Specific Data Length  DWORD   32
'Error Correction Data Length   DWORD   32
'Flags  WORD    16
'Stream Number      7 (LSB)
'Reserved       8
'Encrypted Content Flag     1
'Reserved   DWORD   32
'Type-Specific Data BYTE    varies
'Error Correction Data  BYTE    varies
'
'The fields are defined as follows:
'
'Object ID
'Specifies the GUID for the Stream Properties Object. The value of this field shall be set to ASF_Stream_Properties_Object.
'Object Size
'Specifies the size, in bytes, of the Stream Properties Object. Valid values are at least 78 bytes.
'Stream Type
'Specifies the type of the stream (for example, audio, video, and so on). Use one of the GUIDs defined in section 10.4. Any streams with unrecognized Stream
' Type values should be ignored.
'Error Correction Type
'Specifies the error correction type used by this digital media stream. For streams other than audio, this value should be set to ASF_No_Error_Correction. For
' audio streams, this value should be set to ASF_Audio_Spread. The possible GUIDs are defined in section 10.5, and the standard error correction schemes are
' described in section 9.1.
'Time Offset
'Specifies the presentation time offset of the stream in 100-nanosecond units. The value of this field is added to all of the timestamps of the samples in the
' stream. This value shall be equal to the send time of the first interleaved packet in the data section. The value of this field is typically 0. It is non-zero in the
' case when an ASF file is edited and it is not possible for the editor to change the presentation times and send times of ASF packets. Note that if more than
' one stream is present in an ASF file the offset values of all stream properties objects must be equal.
'Type-Specific Data Length
'Specifies the number of bytes in the Type-Specific Data field.
'Error Correction Data Length
'Specifies the number of bytes in the Error Correction Data field.
'Flags
'The flags are stored in LSB order.
'
'Stream Number (bits 0-6)
'Specifies the number of this stream. 0 is an invalid stream. Valid values are between 1 and 127. The numbers assigned to streams in an ASF presentation
' may be any combination of unique values; parsing logic must not assume that streams are numbered sequentially.
'Reserved (bits 7-14)
'These bits are reserved and should be set to 0.
'Encrypted Content Flag (bit 15)
'Specifies, if set, that the data contained in this stream is encrypted and will be unreadable unless there is a way to decrypt the stream.
'Reserved
'This field is reserved and should be set to 0.
'Type-Specific Data
'Specifies type-specific format data. The structure for the Type-Specific Data field is determined by the value stored in the Stream Type field. The structure for
' the Type-Specific Data field for standard ASF media types is detailed in section 9.
'Error Correction Data
'Specifies data specific to the error correction type. The structure for the Error Correction Data field is determined by the value stored in the Error Correction
' Type field. For example, an audio data stream might need to know how codec chunks were redistributed, or it might need a sample of encoded silence. For
' detailed information, see section 9.1.

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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Dim WholeFile() As Byte
Dim HeaderEnd As Long
Dim Header() As Byte
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
Private Type DlngSize
    Low As Long
    High As Long
End Type
Dim QSize As Currency
Dim CD_QSize As Currency
Private Type GuidStruct
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Dim WMA_GUID As GuidStruct
Dim GUIDIn As GuidStruct
Dim strGUID As String * 36
Dim strHeaderGUID As String * 36
Dim strCD_GUID As String * 36
Dim strInfo_GUID As String * 36
Dim strFile_GUID As String * 36
'
'Dim GUID_Out_0 As Long
'Dim GUID_Out_1 As Long
'Dim GUID_Out_2 As Long
'Dim GUID_Out_3 As Long
Private Type WMAHeaderData
    'each long is 4 bytes...our signature is 16 bytes
    'byte #------------------------------------------------|
    ' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |
    '------------|-------------|-------------|-------------|
    '30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |
    '------------|-------------|-------------|-------------|
    GUID As GuidStruct
    'byte #------------------------------------------>|
    '|17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30 |
    '|------------|-------------|-------------|-------|
    '|51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02 |
    '|------------|-------------|-------------|-------|
    'ObjectSize_Low As Long
    'ObjectSize_High As Long
    ObjectSize As DlngSize
    HeaderObjects As Long 'always 7???
    Reserved1 As Byte 'always 1
    Reserved2 As Byte 'always 2
    '
    'byte #----------------------------------------------->|
    '31-------34 | 35-------38 | 39-------42 | 43-------46 |
    '33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |
    '|-----------|-------------|-------------|-------------|
    CD_GUID As GuidStruct 'Content Description
'    CD_GUID0 As Long
'    CD_GUID1 As Long
'    CD_GUID2 As Long
'    CD_GUID3 As Long
    'byte #-------------------------------------------------------->|
    '|47-------50 | 51-------54 | 55-------58 | 59 60 61 62 | 63 64 |
    '|42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00 00 00 | 00 00 |
    ' ^^                          ^^    ^^                          |
    '                        TitleLen ArtistLen                     |
    '_______________________________________________________________|
    'ObjectSize_Low As Long
    'ObjectSize_High As Long
    CD_ObjectSize As DlngSize
    TitleLen As Integer
    ArtistLen As Integer
    CopyrightLen As Integer
    DescriptionLen As Integer
    RatingLen As Integer
End Type
Dim TagHeader As WMAHeaderData
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
Dim wmaTagCount As Integer
Dim TagIndex As Integer
Dim MaxTags As Integer
Dim i As Integer
Dim j As Integer
Dim fNum As Integer
Dim eM$
Private Type WMAHeaderInfo
    'After the artist's name is the following
    '40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
'    GUID_0 As Long
'    GUID_1 As Long
'    GUID_2 As Long
'    GUID_3 As Long
    GUID As GuidStruct
    '
    'followed by another group that changes
    'This may be a double long to indicate the size of the variable data
    'C2 04 00 00  00 00 00 00 or
    'CE 04 00 00  00 00 00 00 or
    'DC 04 00 00  00 00 00 00
    'HeaderLen As Integer
    'ObjectSize_Low As Long
    'ObjectSize_High As Long
    ObjectSize As DlngSize
    '
    'then i think 2 bytes indicates the number of tags
    TagCount As Integer 'Description Records Count
End Type
Dim TagInfo As WMAHeaderInfo
'This is then followed by some indicators to determine the next data.
'integer for tagname length followed by tagname
'14 00 precedes WM/Lyrics
'12 00 precedes WM/Track
'10 00 is before WM/Year....
'20 00 Encoding
'
'Data Type
'Specifies the type of information being stored. The following values are recognized.
'Value type  Description
'0x0000  Unicode string The data consists of a sequence of Unicode characters.
'0x0001  BYTE array.    The type of data is implementation-specific.
'0x0002  BOOL.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer. Only 0x0000 or 0x0001 are permitted values.
'0x0003  DWORD.         The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'0x0004  QWORD.         The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'0x0005  WORD.          The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.

Dim wmaTagNLen As Integer
Dim wmaTagName() As String
'
Dim wmaTagDType  As Integer
'as near as I can figure, the next 3 bytes after the tag name indicate
'the type of data for the tag.
'0 indicates a string and the length is the next integer
'
Dim wmaTagDLen  As Integer
Dim bytTagData() As Byte
Dim wmaTagData() As Variant
'WM/MCDI, Music CD identifier of the CD from which the file came.
'
Private Type FilePropertiesInfo
    GUID As GuidStruct
    'ObjectSize_Low As Long
    'ObjectSize_High As Long
    ObjectSize As DlngSize
    FileID As GuidStruct
    'FileSize_Low As Long
    'FileSize_High As Long
    FileSize As DlngSize
    'CreationDate_Low As Long
    'CreationDate_High As Long
    CreationDate As DlngSize
    'DataPackets_Low As Long
    'DataPackets_High As Long
    DataPackets As DlngSize
    'PlayDuration_Low As Long
    'PlayDuration_High As Long
    PlayDuration As DlngSize
    'SendDuration_Low As Long
    'SendDuration_High As Long
    SendDuration As DlngSize
    'Preroll_Low As Long
    'Preroll_High As Long
    Preroll As DlngSize
    Flags As Long
    'BroadcastFlag         1 bit (LSB)
    'SeekableFlag          1 bit
    'Reserved             30 bits
    MinPacketSize As Long
    MaxPacketSize As Long
    MaxBitrate As Long
End Type
Dim FileInfo As FilePropertiesInfo
'
Private Type HeaderExtensionInfo
    ObjectID As GuidStruct
    'ObjectSize_Low As Long
    'ObjectSize_High As Long
    ObjectSize As DlngSize 'The value of this field shall be set to 46 bytes  ?
    Reserved1  As GuidStruct 'ASF_Reserved_1
    Reserved2 As Integer 'always 6
    ExtDataSize As Long
End Type
Dim ExtInfo As HeaderExtensionInfo
Dim bytExtData() As Byte
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
'
Private Type CodecListObject
    ObjectID As GuidStruct
    'ObjectSize_Low As Long
    'ObjectSize_High As Long
    ObjectSize As DlngSize 'The value of this field shall be set to 46 bytes
    Reserved  As GuidStruct 'ASF_Reserved_2
    CodecEntriesCount As Long
    'Codec Entries           See below   varies
End Type
Dim CodecListInfo As CodecListObject
Dim CodecCount As Long
'
'Field Name             Field Type  Size (bits)
'----------------------------------------------
'Type                       WORD    16
'Codec Name Length          WORD    16
'Codec Name                 WCHAR  varies
'Codec Description Length   WORD    16
'Codec Description          WCHAR  varies
'Codec Information Length   WORD    16
'Codec Information          BYTE   varies
'
Dim CodecType As Integer    'get the codec type
Dim CodecNameLen As Integer 'get the length of the name
'
Dim bytCodecName() As Byte 'get name data
Dim CodecDLen  As Integer 'Description Length
Dim bytCodecDescr() As Byte 'get descr data
Dim CodecDataLen As Integer 'Codec Data Length
Dim bytCodecData() As Byte 'Codec Information
'below array stores the codec info
Dim CodecName() As String
Dim CodecDescr() As String
Dim CodecData() As Variant
Dim MaxCTags As Integer
Dim CTagIndex As Integer
'
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
Dim bytCodecSpecificData() As Byte 'Codec Information
Dim AudioInfo As AudioCodecInfo
Dim NextGUID As GuidStruct
'
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

Sub ReadWMAHeader(FName As String)
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
Dim myFSO As New FileSystemObject
Dim myTextStream As TextStream
'
Set myTextStream = myFSO.OpenTextFile(FName, ForReading, False)
fstrFullHeader = myTextStream.Read(Val("&H3000"))
myTextStream.Close
'
TagIndex = 0
MaxTags = 0
ReDim wmaTagName(MaxTags)
ReDim wmaTagData(MaxTags)
fNum = FreeFile
Open FName For Binary As fNum
InfoTag.FileSize = LOF(fNum)
'eM$ = "File Length = " & Format(LOF(fNum), "#,###") & " (" & Hex$(LOF(fNum)) & ")" & vbCrLf
fstrInfo = Format(LOF(fNum), "#,###") & " (" & Hex$(LOF(fNum)) & ")" & vbTab
'----------------------------------------------------------------------
'get the beginning main header
Get #fNum, , TagHeader  'this gets the first block of 64 bytes
GUIDIn = TagHeader.GUID
strHeaderGUID = BuildGUID
Debug.Print "   GUID="; strHeaderGUID
Debug.Print Hex$(Seek(fNum)); " After GetTagHeader"
Debug.Print QWORD(TagHeader.ObjectSize.High, TagHeader.ObjectSize.Low)
'now set the len for the title tag
If TagHeader.TitleLen > 0 Then
    ReDim aryTitle(TagHeader.TitleLen - 1)
    Get #fNum, , aryTitle 'this gets the block of bytes for the title
    'wmaTitle = Replace(StrConv(aryTitle, vbUnicode), Chr$(0), "")
    wmaTitle = aryTitle
    wmaTitle = Left$(wmaTitle, Len(wmaTitle) - 1)
    InfoTag.Title = wmaTitle
End If
Debug.Print "Title="; wmaTitle
'
'now set the len for the artist tag
If TagHeader.ArtistLen > 0 Then
    ReDim aryArtist(TagHeader.ArtistLen - 1)
    Get #fNum, , aryArtist 'this gets the block of bytes for the Artist
    'wmaArtist = Replace(StrConv(aryArtist, vbUnicode), Chr$(0), "")
    wmaArtist = aryArtist
    wmaArtist = Left$(wmaArtist, Len(wmaArtist) - 1)
    InfoTag.Artist = wmaArtist
End If
Debug.Print "Artist="; wmaArtist
'
'now set the len for the Copyright tag
If TagHeader.CopyrightLen > 0 Then
    ReDim aryCopyright(TagHeader.CopyrightLen - 1)
    Get #fNum, , aryCopyright 'this gets the block of bytes for the Copyright
    'wmaCopyright = Replace(StrConv(aryCopyright, vbUnicode), Chr$(0), "")
    wmaCopyright = aryCopyright
    wmaCopyright = Left$(wmaCopyright, Len(wmaCopyright) - 1)
    InfoTag.Copyright = wmaCopyright
    InfoTag.IsCopyright = True
Else
    InfoTag.IsCopyright = False
End If
Debug.Print "Copyright="; wmaCopyright
'
'now set the len for the Description tag
If TagHeader.DescriptionLen > 0 Then
    ReDim aryDescription(TagHeader.DescriptionLen - 1)
    Get #fNum, , aryDescription 'this gets the block of bytes for the Description
    'wmaDescription = Replace(StrConv(aryDescription, vbUnicode), Chr$(0), "")
    wmaDescription = aryDescription
    wmaDescription = Left$(wmaDescription, Len(wmaDescription) - 1)
    InfoTag.Description = wmaDescription
End If
Debug.Print "Description="; wmaDescription
'
'now set the len for the Rating tag
If TagHeader.RatingLen > 0 Then
    ReDim aryRating(TagHeader.RatingLen - 1)
    Get #fNum, , aryRating 'this gets the block of bytes for the Rating
    'wmaRating = Replace(StrConv(aryRating, vbUnicode), Chr$(0), "")
    wmaRating = aryRating
    wmaRating = Left$(wmaRating, Len(wmaRating) - 1)
    InfoTag.Rating = wmaRating
End If
Debug.Print "Rating="; wmaRating
Debug.Print Hex$(Seek(fNum)); " After Biblio"
'----------------------------------------------------------------------
'
'byte #------------------------------------------------------------------------------------------------->
' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30 |
'------------|-------------|-------------|-------------|------------|-------------|-------------|-------|
'30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02 |
'            |             |             |             |^^          |             |             |       |------------>|
With TagHeader
    If strHeaderGUID <> ASF_Header_Object Then
        eM$ = eM$ & vbCrLf & vbTab & "TagHeader"
        eM$ = eM$ & vbCrLf & "---------------------------"
        eM$ = eM$ & vbCrLf & "** Error in strHeaderGUID = " & strHeaderGUID
    End If
    If .GUID.Data1 <> 1974609456 Then
        eM$ = eM$ & vbCrLf & vbTab & "TagHeader"
        eM$ = eM$ & vbCrLf & "---------------------------"
        eM$ = eM$ & vbCrLf & "** Error in GUID.Data1 = " & .GUID.Data1 & "(" & Hex$(.GUID.Data1) & ")" '1974609456(75B22630)
    End If
    If .GUID.Data2 <> 26254 Then '298804878 Then changed to integer
        eM$ = eM$ & vbCrLf & "** Error in GUID.Data2 = " & .GUID.Data2 & "(" & Hex$(.GUID.Data2) & ")" '298804878(11CF668E)
    End If
    If .GUID.Data3 <> 4559 Then '-1442784858 Then changed to integer
        eM$ = eM$ & vbCrLf & "** Error in GUID.Data3 = " & .GUID.Data3 & "(" & Hex$(.GUID.Data3) & ")" '-1442784858(AA00D9A6)
    End If
'    If .GUID.Data4 <> 1825464832 Then
'        eM$ = eM$ & vbCrLf & "** Error in GUID.Data4 = " & .GUID.Data4 & "(" & Hex$(.GUID.Data4) & ")" '1825464832(6CCE6200)
'    End If
    'below is always different, so let's show it
    'eM$ = eM$ & vbCrLf & "byte17 = " & .Byte17 & "(" & Hex$(.Byte17) & ")" ' variable
    QSize = QWORD(.ObjectSize.High, .ObjectSize.Low)
    fstrInfo = fstrInfo & QSize & "(" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ")" & vbTab
    If .HeaderObjects <> 7 Then 'always 7?
        eM$ = eM$ & vbCrLf & "** Error in HeaderObjects = " & .HeaderObjects & "(" & Hex$(.HeaderObjects) & ")"
    End If
    If .Reserved1 <> 1 Then 'always 1
        eM$ = eM$ & vbCrLf & "** Error in Reserved1 = " & .Reserved1 & "(" & Hex$(.Reserved1) & ")"
    End If
    If .Reserved2 <> 2 Then 'always 2
        eM$ = eM$ & vbCrLf & "** Error in Reserved2 = " & .Reserved2 & "(" & Hex$(.Reserved2) & ")"
    End If
    '
'31-------34 | 35-------38 | 39-------42 | 43-------46 |47-------50 | 51-------54 | 55-------58 | 59 60 61 62 | 63 64 |
'33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00 00 00 | 00 00 |
'                                                       ^^                          ^^    ^^                          |
'
    GUIDIn = TagHeader.CD_GUID
    strCD_GUID = BuildGUID
    Debug.Print "CD_GUID="; strCD_GUID
    '
    If strCD_GUID <> ASF_Content_Description_Object Then
        eM$ = eM$ & vbCrLf & "** Error in strCD_GUID = " & strCD_GUID
    End If
    If .CD_GUID.Data1 <> 1974609459 Then
        eM$ = eM$ & vbCrLf & "** Error in CD_GUID.Data1 = " & .CD_GUID.Data1 & "(" & Hex$(.CD_GUID.Data1) & ")" '1974609459(75B22633)
    End If
    If .CD_GUID.Data2 <> 26254 Then ' 298804878 Then ' changed to integer
        eM$ = eM$ & vbCrLf & "** Error in CD_GUID.Data2 = " & .CD_GUID.Data2 & "(" & Hex$(.CD_GUID.Data2) & ")" '298804878(11CF668E)
    End If
    If .CD_GUID.Data3 <> 4559 Then ' -1442784858 Then 'changed to integer
        eM$ = eM$ & vbCrLf & "** Error in CD_GUID.Data3 = " & .CD_GUID.Data3 & "(" & Hex$(.CD_GUID.Data3) & ")" '-1442784858(AA00D9A6)
    End If
'    If .CD_GUID.Data4 <> 1825464832 Then
'        eM$ = eM$ & vbCrLf & "** Error in CD_GUID.Data4 = " & .CD_GUID.Data4 & "(" & Hex$(.CD_GUID.Data4) & ")" '1825464832(6CCE6200)
'    End If
    CD_QSize = QWORD(.CD_ObjectSize.High, .CD_ObjectSize.Low)
    fstrInfo = fstrInfo & CD_QSize & "(" & Hex$(.CD_ObjectSize.High) & AddZero(Hex$(.CD_ObjectSize.Low), 8) & ")" & vbTab
    Debug.Print fstrInfo
    '
    'eM$ = eM$ & vbCrLf & "TitleLen = " & .TitleLen & "(" & Hex$(.TitleLen) & ")" ' variable
'    eM$ = eM$ & vbCrLf & wmaTitle
    'eM$ = eM$ & vbCrLf & "ArtistLen = " & .ArtistLen & "(" & Hex$(.ArtistLen) & ")" ' variable
'    eM$ = eM$ & vbCrLf & wmaArtist
    '
    'eM$ = eM$ & vbCrLf & "CopyrightLen = " & .CopyrightLen & "(" & Hex$(.CopyrightLen) & ")" ' variable
'    eM$ = eM$ & vbCrLf & wmaCopyright
    'eM$ = eM$ & vbCrLf & "DescriptionLen = " & .DescriptionLen & "(" & Hex$(.DescriptionLen) & ")" ' variable
'    eM$ = eM$ & vbCrLf & wmaDescription
    'eM$ = eM$ & vbCrLf & "RatingLen = " & .RatingLen & "(" & Hex$(.RatingLen) & ")" ' variable
'    eM$ = eM$ & vbCrLf & wmaRating
End With
If Len(eM$) > 0 Then
    MsgBox eM$
End If
'
Debug.Print Hex$(Seek(fNum)); " Before GetContent"
GetContent
Debug.Print Hex$(Seek(fNum)); " After GetContent"
'
Get #fNum, , FileInfo
GUIDIn = FileInfo.GUID
strFile_GUID = BuildGUID
Debug.Print Hex$(Seek(fNum)); " After File_Properties"
Debug.Print "File GUID="; strFile_GUID
If strGUID = ASF_File_Properties_Object Then
    eM$ = "Executing ASF_File_Properties_Object"
ElseIf strGUID = ASF_Extended_Content_Description_Object Then
    eM$ = "Error!  We received the ASF_Extended_Content_Description_Object"
ElseIf strGUID = ASF_Header_Extension_Object Then
    eM$ = "Error!  We received the ASF_Header_Extension_Object"
ElseIf strGUID = ASF_Codec_List_Object Then
    eM$ = "Error!  We received the ASF_Codec_List_Object"
Else
    eM$ = "Error!  We received another GUID"
End If
MsgBox eM$ & vbCrLf & strGUID
eM$ = ""
If strFile_GUID <> ASF_File_Properties_Object Then
    eM$ = eM$ & vbCrLf & "** Error in strFile_GUID = " & vbCrLf & strFile_GUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
Else
    Debug.Print "ASF_File_Properties_Object"
End If
With FileInfo
    Debug.Print "-------------------ASF_File_Properties_Object-----------------------------"
'    Debug.Print "ObjectSize_Low   = "; .ObjectSize_Low; "("; Hex$(.ObjectSize_Low); ")"
'    Debug.Print "ObjectSize_High  = "; .ObjectSize_High; "("; Hex$(.ObjectSize_High); ")"
    QSize = QWORD(.ObjectSize.High, .ObjectSize.Low)
    Debug.Print "ObjectSize       = "; QSize & "(" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ")" & vbTab
    GUIDIn = .FileID
    strFile_GUID = BuildGUID
    Debug.Print "File GUID="; strFile_GUID
'
'    Debug.Print "FileSize_Low     = "; .FileSize_Low; "("; Hex$(.FileSize_Low); ")"
'    If .FileSize_Low <> InfoTag.FileSize Then
'        eM$ = eM$ & vbCrLf & "** Error in FileSize_Low = " & .FileSize_Low & "(" & Hex$(.FileSize_Low) & ")"
'    End If
'    Debug.Print "FileSize_High    = "; .FileSize_High; "("; Hex$(.FileSize_High); ")"
    Dim FileQSize As Currency
    FileQSize = QWORD(.FileSize.High, .FileSize.Low)
    Debug.Print "FileSize         = "; FileQSize & "(" & Hex$(.FileSize.High) & AddZero(Hex$(.FileSize.Low), 8) & ")" & vbTab
    If FileQSize <> InfoTag.FileSize Then
        eM$ = eM$ & vbCrLf & "** Error in FileSize.Low = " & .FileSize.Low & "(" & Hex$(.FileSize.Low) & ")"
    End If
'
'    Debug.Print "CreationDate_Low     = "; .CreationDate_Low; "("; Hex$(.CreationDate_Low); ")"
'    Debug.Print "CreationDate_High    = "; .CreationDate_High; "("; Hex$(.CreationDate_High); ")"
    'Specifies the date and time of the initial creation of the file.
    'The value is given as the number of 100-nanosecond intervals since January 1, 1601,
    'according to Coordinated Universal Time (Greenwich Mean Time).
    Dim CreationDateQSize As Variant
    CreationDateQSize = QWORD(.CreationDate.High, .CreationDate.Low)
    Debug.Print "CreationDate         = "; Format(CreationDateQSize, "###,#") & "(" & Hex$(.CreationDate.High) & AddZero(Hex$(.CreationDate.Low), 8) & ")" & vbTab
    'Extract date
    Dim DateCreated As Date
    DateCreated = ExtractDate(CreationDateQSize)
    Debug.Print "CreationDate = "; DateCreated
'    Debug.Print "DataPackets_Low  = "; .DataPackets_Low; "("; Hex$(.DataPackets_Low); ")"
'    Debug.Print "DataPackets_High     = "; .DataPackets_High; "("; Hex$(.DataPackets_High); ")"
    Dim DataPacketsQSize As Currency
    DataPacketsQSize = QWORD(.DataPackets.High, .DataPackets.Low)
    Debug.Print "DataPackets         = "; DataPacketsQSize & "(" & Hex$(.DataPackets.High) & AddZero(Hex$(.DataPackets.Low), 8) & ")" & vbTab
'
'    Debug.Print "PlayDuration_Low     = "; .PlayDuration_Low; "("; Hex$(.PlayDuration_Low); ")"
'    Debug.Print "PlayDuration_High    = "; .PlayDuration_High; "("; Hex$(.PlayDuration_High); ")"
    Dim PlayDurationQSize As Currency
    PlayDurationQSize = QWORD(.PlayDuration.High, .PlayDuration.Low)
    Debug.Print "PlayDuration         = "; PlayDurationQSize & "(" & Hex$(.PlayDuration.High) & AddZero(Hex$(.PlayDuration.Low), 8) & ")" & vbTab
    Dim varDuration As Variant
    varDuration = CVar(PlayDurationQSize * ((10 ^ -7)))
    InfoTag.Length = varDuration
    InfoTag.Duration = fTime(InfoTag.Length)
    Debug.Print "Play Duration="; varDuration; " sec.", InfoTag.Duration
'
'    Debug.Print "SendDuration_Low     = "; .SendDuration_Low; "("; Hex$(.SendDuration_Low); ")"
'    Debug.Print "SendDuration_High    = "; .SendDuration_High; "("; Hex$(.SendDuration_High); ")"
    Dim SendDurationQSize As Currency
    SendDurationQSize = QWORD(.SendDuration.High, .SendDuration.Low)
    Debug.Print "sendDuration         = "; SendDurationQSize & "(" & Hex$(.SendDuration.High) & AddZero(Hex$(.SendDuration.Low), 8) & ")" & vbTab
    Dim varSendDuration As Variant
    varSendDuration = CVar(SendDurationQSize * ((10 ^ -7)))
    Debug.Print "Send Duration="; varSendDuration; " sec."
'
'    Debug.Print "Preroll_Low  = "; .Preroll_Low; "("; Hex$(.Preroll_Low); ")"
'    Debug.Print "Preroll_High     = "; .Preroll_High; "("; Hex$(.Preroll_High); ")"
    Dim PreRollQSize As Currency
    PreRollQSize = QWORD(.Preroll.High, .Preroll.Low)
    Debug.Print "PreRoll         = "; PreRollQSize & "(" & Hex$(.Preroll.High) & AddZero(Hex$(.Preroll.Low), 8) & ")" & vbTab
'
    Debug.Print "Flags    = "; .Flags; "("; Hex$(.Flags); ")"
'    Debug.Print "BroadcastFlag1bit(LSB)   = "; .BroadcastFlag1bit(LSB); "("; Hex$(.BroadcastFlag1bit(LSB)); ")"
'    Debug.Print "SeekableFlag1bit     = "; .SeekableFlag1bit; "("; Hex$(.SeekableFlag1bit); ")"
'    Debug.Print "Reserved30bits   = "; .Reserved30bits; "("; Hex$(.Reserved30bits); ")"
    Debug.Print "MinPacketSize    = "; .MinPacketSize; "("; Hex$(.MinPacketSize); ")"
    Debug.Print "MaxPacketSize    = "; .MaxPacketSize; "("; Hex$(.MaxPacketSize); ")"
    Debug.Print "MaxBitrate   = "; .MaxBitrate; "("; Hex$(.MaxBitrate); ")"
End With
'
'========================================================================
'ASF_Header_Extension_Object
Debug.Print Hex$(Seek(fNum)); " Before ExtInfo"
Get #fNum, , ExtInfo
Debug.Print Hex$(Seek(fNum)); " After ExtInfo"
GUIDIn = ExtInfo.ObjectID
strGUID = BuildGUID
Debug.Print "ObjectID="; strGUID
If strGUID = ASF_File_Properties_Object Then
    eM$ = "Error!  We received the ASF_File_Properties_Object"
ElseIf strGUID = ASF_Extended_Content_Description_Object Then
    eM$ = "Error!  We received the ASF_Extended_Content_Description_Object"
ElseIf strGUID = ASF_Header_Extension_Object Then
    eM$ = "Executing ASF_Header_Extension_Object"
ElseIf strGUID = ASF_Codec_List_Object Then
    eM$ = "Error!  We received the ASF_Codec_List_Object"
Else
    eM$ = "Error!  We received another GUID"
End If
Debug.Print eM$
MsgBox eM$ & vbCrLf & strGUID
eM = ""
If strGUID <> ASF_Header_Extension_Object Then
    eM$ = eM$ & vbCrLf & vbTab & "ASF_Header_Extension_Object"
    eM$ = eM$ & vbCrLf & "---------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strGUID = " & vbCrLf & strGUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
GUIDIn = ExtInfo.Reserved1
strGUID = BuildGUID
If strGUID <> ASF_Reserved_1 Then
    eM$ = eM$ & vbCrLf & "** Error in ASF_Reserved_1 = " & strGUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
'
With ExtInfo
    Debug.Print "-------------------ASF_Header_Extension_Object-----------------------------"
'    Debug.Print "ObjectSize_Low   = "; .ObjectSize_Low; "("; Hex$(.ObjectSize_Low); ")"
'    Debug.Print "ObjectSize_High  = "; .ObjectSize_High; "("; Hex$(.ObjectSize_High); ")"
    QSize = QWORD(.ObjectSize.High, .ObjectSize.Low)
    Debug.Print "ObjectSize       = "; QSize & "(" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ")" & vbTab
    GUIDIn = .Reserved1
    strGUID = BuildGUID
    Debug.Print "Reserved1 GUID="; strGUID
    Debug.Print "Reserved2="; .Reserved2
    '
    'Specifies the number of bytes stored in the Header Extension Data field.
    'This value may be 0 bytes or 24 bytes and larger.
    'It should also be equal to the Object Size field minus 46 bytes.
    Debug.Print "ExtDataSize   = "; .ExtDataSize; "("; Hex$(.ExtDataSize); ")"
    ReDim bytExtData(.ExtDataSize - 1)
    If .ExtDataSize <> QSize - 46 Then
        eM$ = "There may be an error with the size of the extended data..."
        eM$ = eM$ & vbCrLf & "ObjectSize    = " & QSize & "(" & Hex$(QSize) & ")"
        eM$ = eM$ & vbCrLf & "ExtDataSize   = " & .ExtDataSize & "(" & Hex$(.ExtDataSize) & ")"
    End If
End With
'let's just get the data like we would normally
'GoTo ExtDataDone
'-----------------------------------------------------------
Get #fNum, , bytExtData
'move the data to a string
Dim ExtensionData As String
ExtensionData = bytExtData
ExtensionData = Left$(ExtensionData, Len(ExtensionData) - 1)
'Debug.Print ExtensionData
'
Dim HexStr$
'Debug.Print AddZero(Hex$((0)), 6); vbTab;
For i = 0 To UBound(bytExtData) - 1
    If i Mod 16 = 0 Then
        Debug.Print HexStr$
        HexStr$ = ""
        Debug.Print AddZero(Hex$((i)), 6); vbTab;
    End If
    HexStr$ = HexStr$ & AddZero(Hex$(bytExtData(i)), 2) & " "
    If i > 160 Then
        If HexStr$ = "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " Then
        'If bytExtData(i - 2) = 0 And bytExtData(i - 1) = 0 And bytExtData(i) = 0 Then
            Exit For
        End If
    End If
Next i
Debug.Print
Debug.Print Hex$(Seek(fNum)); " After bytExtData"
'=========================================================================
ExtDataDone:
'ASF_Codec_List_Object
Debug.Print Hex$(Seek(fNum)); " Before CodecListInfo"
Get #fNum, , CodecListInfo
Debug.Print Hex$(Seek(fNum)); " After CodecListInfo"
GUIDIn = CodecListInfo.ObjectID
strGUID = BuildGUID
'should be 86D15240-311D-11D0-A3A4-00A0C90348F6
Debug.Print "ObjectID="; strGUID
If strGUID = ASF_File_Properties_Object Then
    eM$ = "Error!  We received the ASF_File_Properties_Object"
ElseIf strGUID = ASF_Extended_Content_Description_Object Then
    eM$ = "Error!  We received the ASF_Extended_Content_Description_Object"
ElseIf strGUID = ASF_Header_Extension_Object Then
    eM$ = "Error!  We received the ASF_Header_Extension_Object"
ElseIf strGUID = ASF_Codec_List_Object Then
    eM$ = "Executing ASF_Codec_List_Object"
'    GoTo ASF_Codec_List_Object
ElseIf strGUID = ASF_Audio_Media Then
    eM$ = "Error!  We received the ASF_Audio_Media"
'    GoTo ASF_Audio_Media
Else
    eM$ = "Error!  We received another GUID"
End If
ASF_Codec_List_Object:
MsgBox eM$ & vbCrLf & strGUID
eM$ = ""
If strGUID <> ASF_Codec_List_Object Then
    eM$ = eM$ & vbCrLf & vbTab & "ASF_Codec_List_Object"
    eM$ = eM$ & vbCrLf & "---------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strGUID = " & vbCrLf & strGUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
'-------------------------------------------------------------------------------------
With CodecListInfo
    Debug.Print "-------------------ASF_Codec_List_Object-----------------------------"
'    Debug.Print "ObjectSize_Low   = "; .ObjectSize_Low; "("; Hex$(.ObjectSize_Low); ")"
'    Debug.Print "ObjectSize_High  = "; .ObjectSize_High; "("; Hex$(.ObjectSize_High); ")"
    QSize = QWORD(.ObjectSize.High, .ObjectSize.Low)
    Debug.Print "ObjectSize       = "; QSize & "(" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ")" & vbTab
    '
    GUIDIn = .Reserved
    'should be 86D15241-311D-11D0-A3A4-00A0C90348F6
    strGUID = BuildGUID
    If strGUID <> ASF_Reserved_2 Then
        eM$ = eM$ & vbCrLf & "** Error in ASF_Reserved_2 = " & strGUID
    End If
    Debug.Print "Reserved GUID="; strGUID
    Debug.Print "CodecEntriesCount  = "; .CodecEntriesCount; "("; Hex$(.CodecEntriesCount); ")"
End With
Debug.Print Hex$(Seek(fNum)); " After CodecListInfo"
GetCodecs
'
'-------------------------------------------------------------------------------
ASF_Audio_Media:
'we must first get ASF_Stream_Properties_Object
'this should be B7DC0791-A9B7-11CF-8EE6-00C00C205365
'then we dig out the ASF_Audio_Media
'F8699E40-5B4D-11CF-A8FD-00805F5C442B
Get #fNum, , NextGUID
GUIDIn = NextGUID
strGUID = BuildGUID
Debug.Print "GUID="; strGUID
'
If strGUID = ASF_File_Properties_Object Then
    eM$ = "Error!  We received the ASF_File_Properties_Object"
ElseIf strGUID = ASF_Extended_Content_Description_Object Then
    eM$ = "Error!  We received the ASF_Extended_Content_Description_Object"
ElseIf strGUID = ASF_Header_Extension_Object Then
    eM$ = "Error!  We received the ASF_Header_Extension_Object"
ElseIf strGUID = ASF_Codec_List_Object Then
    eM$ = "Error!  We received the ASF_Codec_List_Object"
'    GoTo ASF_Codec_List_Object
ElseIf strGUID = ASF_Audio_Media Then
    eM$ = "Error!  We received the ASF_Audio_Media"
ElseIf strGUID = ASF_Stream_Properties_Object Then
    eM$ = "Executing ASF_Stream_Properties_Object"
Else
    eM$ = "Error!  We received another GUID"
End If
MsgBox eM$ & vbCrLf & strGUID
eM$ = ""
If strGUID <> ASF_Stream_Properties_Object Then
    eM$ = eM$ & vbCrLf & vbTab & "ASF_Stream_Properties_Object"
    eM$ = eM$ & vbCrLf & "---------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strGUID = " & vbCrLf & strGUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
'
If strGUID <> ASF_Audio_Media Then
    eM$ = eM$ & vbCrLf & vbTab & "ASF_Audio_Media"
    eM$ = eM$ & vbCrLf & "---------------------------"
    eM$ = eM$ & vbCrLf & "** Error in strGUID = " & vbCrLf & strGUID
    Debug.Print eM$
    MsgBox eM$
    eM$ = ""
End If
'
Get #fNum, , AudioInfo 'ASF_Audio_Media
With AudioInfo
    Debug.Print "-------------------ASF_Audio_Media-----------------------------"
    Debug.Print "CodecID   = "; .CodecID; "("; Hex$(.CodecID); ")"
    Debug.Print "NumChannels  = "; .NumChannels; "("; Hex$(.NumChannels); ")"
    Debug.Print "SamplesPerSec   = "; .SamplesPerSec; "("; Hex$(.SamplesPerSec); ")"
    Debug.Print "AvgBytesPerSec   = "; .AvgBytesPerSec; "("; Hex$(.AvgBytesPerSec); ")"
    Debug.Print "BlockAlignment   = "; .BlockAlignment; "("; Hex$(.BlockAlignment); ")"
    Debug.Print "BitsPerSample   = "; .BitsPerSample; "("; Hex$(.BitsPerSample); ")"
    Debug.Print "CodecDataSize   = "; .CodecDataSize; "("; Hex$(.CodecDataSize); ")"
End With
Debug.Print Hex$(Seek(fNum)); " After AudioInfo"
'-------------------------------------------------------------------------------
'
'===============================================================================
AllDone:
'
'-------------------------------------------------------------------------------
GetGUID
'
Close #fNum
If Len(eM$) > 0 Then
    MsgBox eM$
End If
Exit Sub
'===============================================================================
'
'You can determine how the data in a Variant is treated using the
'VarType function or TypeName function
'From VarType
'vbEmpty 0 Empty (uninitialized)
'vbNull 1 Null (no valid data)
'vbInteger 2 Integer
'vbLong 3 Long integer
'vbSingle 4 Single-precision floating-point number
'vbDouble 5 Double-precision floating-point number
'vbCurrency 6 Currency value
'vbDate 7 Date value
'vbString 8 String
'vbObject 9 Object
'vbError 10 Error value
'vbBoolean 11 Boolean value
'vbVariant 12 Variant (used only witharrays of variants)
'vbDataObject 13 A data access object
'vbDecimal 14 Decimal value
'vbByte 17 Byte value
'vbUserDefinedType 36 Variants that contain user-defined types
'vbArray 8192 Array
'
'after the last tag, is:
'A1 DC AB 8C  47 A9 CF 11  8E E4 00 C0  0C 20 53 65  68 00 00 00
'this is good becuase the temp files saved after listening to streamed music starts with
'68 26 B2 75  8E 66 CF 11  A6 D9 00 AA  00 62 CE 6C
'3D 13 00 00  00 00 00 00  07 00 00 00  01 02 Then at 1EH (30)...
'A1 DC  AB 8C 47 A9  CF 11 8E E4  00 C0 0C 20  53 65 68 00
'our magic header
'so maybe if i build the header and paste it before above, I can play the temp file?
HeaderEnd = Seek(fNum)
eM$ = ""
Dim lngTemp As Long
Get #fNum, , lngTemp
If lngTemp <> Val("&H8CABDCA1") Then
    eM$ = eM$ & vbCrLf & "** Error in lngTemp = " & lngTemp & "(" & Hex$(lngTemp) & ")"
End If
Get #fNum, , lngTemp
If lngTemp <> Val("&H11CFA947") Then
    eM$ = eM$ & vbCrLf & "** Error in lngTemp = " & lngTemp & "(" & Hex$(lngTemp) & ")"
End If
Get #fNum, , lngTemp
If lngTemp <> Val("&HC000E48E") Then
    eM$ = eM$ & vbCrLf & "** Error in lngTemp = " & lngTemp & "(" & Hex$(lngTemp) & ")"
End If
Get #fNum, , lngTemp
If lngTemp <> Val("&H6553200C") Then
    eM$ = eM$ & vbCrLf & "** Error in lngTemp = " & lngTemp & "(" & Hex$(lngTemp) & ")"
End If
Get #fNum, , lngTemp
If lngTemp <> Val("&H68") Then
    eM$ = eM$ & vbCrLf & "** Error in lngTemp = " & lngTemp & "(" & Hex$(lngTemp) & ")"
End If
Close fNum
If Len(eM$) > 0 Then
    MsgBox eM$
    eM$ = ""
End If
'now reopen and get just the header
HeaderEnd = Val("&H6EF")
ReDim Header(HeaderEnd)
fNum = FreeFile
Open FName For Binary As fNum
Get #fNum, , Header 'this gets the whole header
Close #fNum
Exit Sub
ReDim WholeFile(LOF(fNum))
'now open our temp file
fNum = FreeFile
Open "C:\Documents and Settings\pydeskto\Desktop\New Folder\Lithium.wma" For Binary As fNum
ReDim WholeFile(LOF(fNum) - 480) '1eh replaces the whole header
Get #fNum, 481, WholeFile  'this gets the whole rest of the file
Close #fNum
'now build a new file
Open "C:\Documents and Settings\pydeskto\Desktop\New Folder\Lithium2.wma" For Binary As fNum
Put #fNum, , Header
Put #fNum, HeaderEnd, WholeFile
Close #fNum
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

Sub GetContent()
On Error GoTo Oops
Get #fNum, , TagInfo 'this gets a block of 26 bytes after the title/artist info
'
eM$ = ""
With TagInfo
    'After the artist's name is the following
    '40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
    'which is the ASF_Extended_Content_Description_Object
    GUIDIn = .GUID
    strInfo_GUID = BuildGUID
    Debug.Print "strInfo_GUID= "; strInfo_GUID
    If strInfo_GUID <> ASF_Extended_Content_Description_Object Then
        eM$ = eM$ & vbCrLf & "___________________________"
        eM$ = eM$ & vbCrLf & vbCrLf & vbTab & "TagInfo"
        eM$ = eM$ & vbCrLf & "---------------------------"
        eM$ = eM$ & vbCrLf & "** Error in strInfo_GUID = " & strInfo_GUID
    End If
    If .GUID.Data1 <> -758078400 Then
        eM$ = eM$ & vbCrLf & "GUID.Data1 = " & .GUID.Data1 & "(" & Hex$(.GUID.Data1) & ")"            '-758078400(D2D0A440)
    End If
    If .GUID.Data2 <> -7417 Then ' 299033351 Then 'changed to integer
        eM$ = eM$ & vbCrLf & "** Error in GUID.Data2 = " & .GUID.Data2 & "(" & Hex$(.GUID.Data2) & ")" '299033351(11D2E307)
    End If
    If .GUID.Data3 <> 4562 Then ' -1610551145 Then 'changed to integer
        eM$ = eM$ & vbCrLf & "** Error in GUID.Data3 = " & .GUID.Data3 & "(" & Hex$(.GUID.Data3) & ")" '-1610551145(A000F097)
    End If
'    If .GUID.Data4 <> 1353211593 Then
'        eM$ = eM$ & vbCrLf & "** Error in GUID.Data4 = " & .GUID.Data4 & "(" & Hex$(.GUID.Data4) & ")" '1353211593(50A85EC9)
'    End If
    '
    'followed by another group that changes
    'C2 04 00 00  00 00 00 00 or
'    eM$ = eM$ & vbCrLf & "byte17 = " & .Byte17 & "(" & Hex$(.Byte17) & ")" 'variable
    QSize = QWORD(.ObjectSize.High, .ObjectSize.Low)
    Debug.Print "ObjectSize       = "; QSize & "(" & Hex$(.ObjectSize.High) & AddZero(Hex$(.ObjectSize.Low), 8) & ")" & vbTab
    fstrInfo = fstrInfo & QSize & "(" & Hex$(QSize) & ")" & vbTab
    '
    'then i think 2 bytes indicates the number of tags
    'This is then followed by some indicators to determine the next data.
'    eM$ = eM$ & vbCrLf & "TagCount = " & .TagCount & "(" & Hex$(.TagCount) & ")"
    wmaTagCount = .TagCount
End With
Debug.Print fstrInfo
Debug.Print eM$
If Len(eM$) Then MsgBox eM$, , "TagHeader Info " '& wmaTitle
'
ReDim Preserve wmaTagName(MaxTags + wmaTagCount)
ReDim Preserve wmaTagData(MaxTags + wmaTagCount)
eM$ = ""
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
For i = MaxTags To wmaTagCount - 1
    Get #fNum, , wmaTagNLen 'Name Length
    ReDim bytTagData(wmaTagNLen - 1)
    'get the name of the tag
    Get #fNum, , bytTagData
    TagIndex = MaxTags + i
    'convert the byte data to a string by forcing it into the string value
    wmaTagName(TagIndex) = bytTagData
    'trim the null from the enc
    wmaTagName(TagIndex) = Left$(wmaTagName(TagIndex), Len(wmaTagName(TagIndex)) - 1)
    'get the Data Type of the tag
    Get #fNum, , wmaTagDType  '0 indicates a string
    'get the length of the Tag Data
    Get #fNum, , wmaTagDLen
    'Debug.Print "wmaTagNLen="; wmaTagNLen; "("; Hex$(wmaTagNLen); ")"; " - ";
    'Debug.Print "wmaTagDType="; wmaTagDType; " - ";
    'Debug.Print "wmaTagDLen="; wmaTagDLen; "("; Hex$(wmaTagDLen); ")"
    ReDim bytTagData(wmaTagDLen - 1)
    'get the tag data
    Get #fNum, , bytTagData
'0x0000  Unicode string The data consists of a sequence of Unicode characters.
'0x0001  BYTE array The type of data is implementation-specific.
'0x0002  BOOL.      The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer. Only 0x0000 or 0x0001 are permitted values.
'0x0003  DWORD.     The data is 4 bytes long and should be interpreted as a 32-bit unsigned integer.
'0x0004  QWORD.     The data is 8 bytes long and should be interpreted as a 64-bit unsigned integer.
'0x0005  WORD.      The data is 2 bytes long and should be interpreted as a 16-bit unsigned integer.
'
    'the Tag data must be set according to its datatype
    Select Case wmaTagDType
        Case Is = 0 'string values
            wmaTagData(TagIndex) = bytTagData
            wmaTagData(TagIndex) = Left$(wmaTagData(TagIndex), Len(wmaTagData(TagIndex)) - 1)
        Case Is = 1  'BYTE array
            wmaTagData(TagIndex) = bytTagData
            wmaTagData(TagIndex) = Left$(wmaTagData(TagIndex), Len(wmaTagData(TagIndex)) - 1)
        Case 2 'Single  !!Supposed to be Boolean!!  VB boolean is 2 bytes - MetaData is 4 bytes
            Dim tmpSingle As Single
            CopyMemory ByVal VarPtr(tmpSingle), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(TagIndex) = CBool(tmpSingle)
        Case 3 'Long
            Dim tmpLong As Long
            CopyMemory ByVal VarPtr(tmpLong), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(TagIndex) = tmpLong
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
            CopyMemory ByVal VarPtr(tmpDouble), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            CopyMemory ByVal VarPtr(tmpDate), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen / 2)
            CopyMemory ByVal VarPtr(tmpTime), ByVal VarPtr(bytTagData(wmaTagDLen / 2)), CLng(wmaTagDLen / 2)
            wmaTagData(TagIndex) = tmpDouble
            Dim HexStr$
            For j = 0 To wmaTagDLen - 1
                HexStr$ = Right$("0" & (Hex$(bytTagData(j))), 2) & HexStr$
            Next j
            tmpDouble = CDbl(Val("&H" & HexStr$))
            tmpDate = CLng("&H" & Left(HexStr$, 8))
            tmpTime = CLng("&H" & Right(HexStr$, 8))
            'tmpD = CDate(tmpDouble) 'CLng("&H" & Left(hexstr$, 8)))
        Case 5 'integer
            Dim tmpInteger As Integer
            CopyMemory ByVal VarPtr(tmpInteger), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(TagIndex) = tmpInteger
        Case Else
            Dim wmaVariant As Variant
            CopyMemory (wmaVariant), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(TagIndex) = wmaVariant
    End Select
    eM$ = eM$ & wmaTagName(TagIndex) & " = " & wmaTagData(TagIndex) & vbCrLf
'    Debug.Print wmaTagName(TagIndex); " = "; wmaTagData(TagIndex),
'    Debug.Print VarType(wmaTagData(TagIndex)); " "; TypeName(wmaTagData(TagIndex))
    With InfoTag
        Select Case wmaTagName(TagIndex)
            Case "WM/Track"
            Case "WM/Lyrics"
            Case "WM/MediaPrimaryClassID"
            Case "WMFSDKVersion"
            Case "WMFSDKNeeded"
            Case "IsVBR"
            Case "WM/TrackNumber"
                .TrackNumber = wmaTagData(TagIndex)
            Case "WM/Year"
                .Year = wmaTagData(TagIndex)
            Case "WM/EncodingTime"
            Case "WM/UniqueFileIdentifier"
            Case "WM/Composer"
                .Composer = wmaTagData(TagIndex)
            Case "WM/Publisher"
                .Publisher = wmaTagData(TagIndex)
            Case "WM/Genre"
                .Genre = wmaTagData(TagIndex)
            Case "WM/AlbumTitle"
                .Album = wmaTagData(TagIndex)
            Case "WM/AlbumArtist"
                .AlbumArtist = wmaTagData(TagIndex)
            Case "WM/MCDI"
            Case "WM/Provider"
            Case "WM/ProviderRating"
                .Rating = wmaTagData(TagIndex)
            Case "WM/ProviderStyle"
        End Select
    End With
Next i
MaxTags = MaxTags + wmaTagCount
'
Debug.Print "----------------------------------------------------------"
Debug.Print eM$
Debug.Print "----------------------------------------------------------"
'MsgBox eM$, , "Tag Data - Tag Count = " & wmaTagCount
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
GoTo Exit_GetContent
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetContent "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GetContent"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetContent:
End Sub

Sub GetCodecs()
On Error GoTo Oops
'
'Get #fNum, , CodecListInfo
'If strGUID <> ASF_Codec_List_Object Then
'    eM$ = eM$ & vbCrLf & vbTab & "ASF_Codec_List_Object"
'    eM$ = eM$ & vbCrLf & "---------------------------"
'    eM$ = eM$ & vbCrLf & "** Error in strGUID = " & strGUID
'End If
'With CodecListInfo
'    Debug.Print "-------------------ASF_Codec_List_Object-----------------------------"
'    Debug.Print "ObjectSize_Low   = "; .ObjectSize_Low; "("; Hex$(.ObjectSize_Low); ")"
'    Debug.Print "ObjectSize_High  = "; .ObjectSize_High; "("; Hex$(.ObjectSize_High); ")"
'    GUIDIn = .Reserved1
'    strGUID = BuildGUID
'    Debug.Print "Reserved1 GUID="; strGUID
'    Debug.Print "CodecEntriesCount  = "; .CodecEntriesCount; "("; Hex$(.CodecEntriesCount); ")"
'End With
'Debug.Print Hex$(Seek(fNum)); " After CodecListInfo"
'
CodecCount = CodecListInfo.CodecEntriesCount
ReDim Preserve CodecName(CodecCount)
ReDim Preserve CodecDescr(CodecCount)
ReDim Preserve CodecData(CodecCount)
Debug.Print "Retrieving data for "; CodecCount; " Codec Entries."
eM$ = ""
'
For i = MaxCTags To CodecCount - 1
    Get #fNum, , CodecType 'Name Length
    If CodecType <> 2 Then
        MsgBox "This is not an audio codec!"
    End If
    Debug.Print "CodecType="; CodecType
    'get Length of the name data
    Get #fNum, , CodecNameLen
    Debug.Print "CodecNameLen="; CodecNameLen
    ReDim bytCodecName(119) ' (CodecNameLen - 1)
    'get the name of the CTag
    Get #fNum, , bytCodecName
    CTagIndex = MaxCTags + i
    'convert the byte data to a string by forcing it into the string value
    CodecName(CTagIndex) = bytCodecName
    'trim the null from the Name
    'CodecName(CTagIndex) = Left$(CodecName(CTagIndex), Len(CodecName(CTagIndex)) - 1)
    Debug.Print CodecName(CTagIndex)
    '
    'get the Description length
    Get #fNum, , CodecDLen
    If CodecDLen > 0 Then
        ReDim bytCodecDescr(CodecDLen - 1)
        'get the Descr of the CTag
        Get #fNum, , bytCodecDescr
        'convert the byte data to a string by forcing it into the string value
        CodecDescr(CTagIndex) = bytCodecDescr
        'trim the null from the Descr
        CodecDescr(CTagIndex) = Left$(CodecDescr(CTagIndex), Len(CodecDescr(CTagIndex)) - 1)
    End If
    '
    'Get the Codec Data
    Get #fNum, , CodecDataLen
    If CodecDataLen > 0 Then
        ReDim bytCodecData(CodecDataLen - 1)
        'get the Data of the CTag
        Get #fNum, , bytCodecData
        'convert the byte data to a string by forcing it into the string value
        CodecData(CTagIndex) = bytCodecData
        'trim the null from the Data
        'CodecData(CTagIndex) = Left$(CodecData(CTagIndex), Len(CodecData(CTagIndex)) - 1)
    End If
    '
    eM$ = eM$ & CodecName(CTagIndex) & " = " & CodecDescr(CTagIndex) & " = " & CodecData(CTagIndex) & vbCrLf
    With InfoTag
        Select Case CodecName(CTagIndex)
            Case "WM/TrackNumber"
                .TrackNumber = CodecData(CTagIndex)
        End Select
    End With
Next i
MaxCTags = MaxCTags + CodecCount
'
Debug.Print "----------------------------------------------------------"
Debug.Print eM$
Debug.Print "----------------------------------------------------------"
'MsgBox eM$, , "CTag Data - CTag Count = " & wmaCTagCount
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
End Sub

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
    Debug.Print "strTrack="; Asc(strTrack), strTrack
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
Dim intRemMin As Integer, strRemMin As String
Dim strMin As String, intMin As Integer, intSec As Integer
Dim strHour As String, intHour As Integer
strHour = CStr((lngLength / 360))
If InStr(1, strHour, ".") > 0 Then
    strHour = Left(strHour, InStr(1, strHour, ".") - 1)
End If
If Not strHour = "" Then
    intHour = CInt(strHour)
Else
    intHour = 0
End If
strRemMin = CStr(CDbl(lngLength Mod 360))
If InStr(1, strRemMin, ".") > 0 Then
    strRemMin = Left(strRemMin, InStr(1, strRemMin, ".") - 1)
End If
If Not strRemMin = "" Then
    intRemMin = CInt(strRemMin)
Else
    intRemMin = 0
End If
strMin = CStr((intRemMin / 60))
If InStr(1, strMin, ".") > 0 Then
    strMin = Left(strMin, InStr(1, strMin, ".") - 1)
End If
If Not strMin = "" Then
    intMin = CInt(strMin)
Else
    intMin = 0
End If
intSec = CInt(lngLength Mod 60)
fTime = Double_digits(intHour) & ":" & Double_digits(intMin) & ":" & Double_digits(intSec) & ""
End Function

Function Double_digits(intInput As Integer) As String
Dim strInput As String
strInput = CStr(intInput)
If Len(strInput) < 2 Then
    strInput = "0" & strInput
End If
Double_digits = strInput
End Function

Sub PutWMATag(FName As String)
On Error GoTo Oops
Dim i As Integer
Dim j As Integer
Dim fNum As Integer
Dim eM$
'
fNum = FreeFile
'
Open FName For Binary As fNum
'
'byte #------------------------------------------------------------------------------------------------->
' 1  2  3  4 |  5  6  7  8 | 9  10 11 12 | 13 14 15 16 |17 18 19 20 | 21 22 23 24 | 25 26 27 28 | 29 30 |
'------------|-------------|-------------|-------------|------------|-------------|-------------|-------|
'30 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |51 16 00 00 | 00 00 00 00 | 07 00 00 00 | 01 02 |
'            |             |             |             |^^          |             |             |       |------------>|
With TagHeader
'    .GUID_0 = 1974609456
'    .GUID_1 = 298804878
'    .GUID_2 = -1442784858
'    .GUID_3 = 1825464832
    .GUID.Data1 = 1974609456
    .GUID.Data2 = 26254 '298804878  changed to integer
    .GUID.Data3 = 4559  '-1442784858 changed to integer
'    .GUID.Data4 = 1825464832
    'below is always different, so let's show it
.ObjectSize.Low = 0 ' variable
    .ObjectSize.High = 0
    .HeaderObjects = 7
    .Reserved1 = 1
    .Reserved2 = 2
    '
'31-------34 | 35-------38 | 39-------42 | 43-------46 |47-------50 | 51-------54 | 55-------58 | 59 60 61 62 | 63 64 |
'33 26 B2 75 | 8E 66 CF 11 | A6 D9 00 AA | 00 62 CE 6C |42 00 00 00 | 00 00 00 00 | 18 00 08 00 | 00 00 00 00 | 00 00 |
'                                                       ^^                          ^^    ^^                          |
'                                                                               TitleLen ArtistLen                    |
    .CD_GUID.Data1 = 1974609459
    .CD_GUID.Data2 = 26254  ' 298804878 changed to integer
    .CD_GUID.Data3 = 4559  ' -1442784858 changed to integer
    '.CD_GUID.Data4 = 1825464832
.CD_ObjectSize.Low = 0 ' variable
    .CD_ObjectSize.High = 0
    '
    eM$ = eM$ & vbCrLf & "TitleLen = " & .TitleLen & "(" & Hex$(.TitleLen) & ")" ' variable
    eM$ = eM$ & vbCrLf & wmaTitle
    eM$ = eM$ & vbCrLf & "ArtistLen = " & .ArtistLen & "(" & Hex$(.ArtistLen) & ")" ' variable
    eM$ = eM$ & vbCrLf & wmaArtist
    '
    .CopyrightLen = 0
    .DescriptionLen = 0
    .RatingLen = 0
End With
'
Put #fNum, , TagHeader
'now set the len for the title and artist tags
ReDim aryTitle(TagHeader.TitleLen - 1)
aryTitle = StrConv(wmaTitle & Chr$(0), vbFromUnicode)
Put #fNum, , aryTitle 'this Puts the block of bytes for the title
ReDim aryArtist(TagHeader.ArtistLen - 1)
aryArtist = StrConv(wmaArtist & Chr$(0), vbFromUnicode)
Put #fNum, , aryArtist 'this Puts the block of bytes for the Artist
'
With TagInfo
    eM$ = eM$ & vbCrLf & "___________________________"
    eM$ = eM$ & vbCrLf & vbCrLf & vbTab & "TagInfo"
    eM$ = eM$ & vbCrLf & "---------------------------"
    'After the artist's name is the following
    '40 A4 D0 D2  07 E3 D2 11  97 F0 00 A0  C9 5E A8 50
    .GUID.Data1 = -758078400
    .GUID.Data2 = -7417  ' 299033351 changed to integer
    .GUID.Data3 = 4562  ' -1610551145 changed to integer
    '.GUID.Data4 = 1353211593
    '
    'followed by another group that changes
    'C2 04 00 00  00 00 00 00 or
.ObjectSize.Low = 0 'variable
    .ObjectSize.High = 0
    .TagCount = wmaTagCount
End With
Put #fNum, , TagInfo
'
ReDim wmaTagName(TagInfo.TagCount)
ReDim wmaTagData(TagInfo.TagCount)
eM$ = ""
For i = 1 To TagInfo.TagCount
    Put #fNum, , wmaTagNLen
    ReDim bytTagData(wmaTagNLen - 1)
    Put #fNum, , bytTagData
    wmaTagName(i) = bytTagData
    wmaTagName(i) = Left$(wmaTagName(i), Len(wmaTagName(i)) - 1)
    Put #fNum, , wmaTagDType
    '0 indicates a string
    Put #fNum, , wmaTagDLen
    'Debug.Print "wmaTagNLen="; wmaTagNLen; "("; Hex$(wmaTagNLen); ")"; " - ";
    'Debug.Print "wmaTagDType="; wmaTagDType; " - ";
    'Debug.Print "wmaTagDLen="; wmaTagDLen; "("; Hex$(wmaTagDLen); ")"
    ReDim bytTagData(wmaTagDLen - 1)
    Put #fNum, , bytTagData
    Select Case wmaTagDType
        Case Is = 0, 1 'string values
            wmaTagData(i) = bytTagData
            wmaTagData(i) = Left$(wmaTagData(i), Len(wmaTagData(i)) - 1)
        Case 2 'Single
            Dim tmpSingle As Single
            CopyMemory ByVal VarPtr(tmpSingle), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(i) = tmpSingle
        Case 3 'Long
            Dim tmpLong As Long
            CopyMemory ByVal VarPtr(tmpLong), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(i) = tmpLong
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
            CopyMemory ByVal VarPtr(tmpDouble), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            CopyMemory ByVal VarPtr(tmpDate), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen / 2)
            CopyMemory ByVal VarPtr(tmpTime), ByVal VarPtr(bytTagData(wmaTagDLen / 2)), CLng(wmaTagDLen / 2)
            wmaTagData(i) = tmpDouble
            Dim HexStr$
            For j = 0 To wmaTagDLen - 1
                HexStr$ = Right$("0" & (Hex$(bytTagData(j))), 2) & HexStr$
            Next j
            tmpDouble = CDbl(Val("&H" & HexStr$))
            tmpDate = CLng("&H" & Left(HexStr$, 8))
            tmpTime = CLng("&H" & Right(HexStr$, 8))
            'tmpD = CDate(tmpDouble) 'CLng("&H" & Left(hexstr$, 8)))
        Case Else
            Dim wmaVariant As Variant
            CopyMemory (wmaVariant), ByVal VarPtr(bytTagData(0)), CLng(wmaTagDLen)
            wmaTagData(i) = wmaVariant
    End Select
    eM$ = eM$ & wmaTagName(i) & " = " & wmaTagData(i) & vbCrLf
'    Debug.Print wmaTagName(i); " = "; wmaTagData(i),
'    Debug.Print VarType(wmaTagData(i)); " "; TypeName(wmaTagData(i))
Next i
'
Close fNum
'
GoTo Exit_PutWMATag
Oops:
'Abort=3,Retry=4,IeTitle$ = App.Title & ": Error in Subroutine PutWMATag "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in PutWMATag"
eMess$ = eMess$ & IIf(Erl = 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_PutWMATag:
End Sub

Sub GetGUID()
Get #fNum, , NextGUID
GUIDIn = NextGUID
strGUID = BuildGUID
Debug.Print "GUID="; strGUID
If strGUID = ASF_File_Properties_Object Then
    eM$ = "ASF_File_Properties_Object"
ElseIf strGUID = ASF_Extended_Content_Description_Object Then
    eM$ = "ASF_Extended_Content_Description_Object"
ElseIf strGUID = ASF_Header_Extension_Object Then
    eM$ = "ASF_Header_Extension_Object"
ElseIf strGUID = ASF_Codec_List_Object Then
    eM$ = "ASF_Codec_List_Object"
ElseIf strGUID = ASF_Audio_Media Then
    eM$ = "ASF_Audio_Media"
Else
    eM$ = "Other GUID"
End If
MsgBox eM$ & vbCrLf & strGUID
eM$ = ""
Debug.Print Hex$(Seek(fNum)); " After guid"
End Sub

Function BuildGUID() As String
'make it look like 75B22630-668E-11CF-A6D9-00AA0062CE6C
Dim byt As Integer
With GUIDIn
    BuildGUID = AddZero(Hex$(.Data1), 8) & "-"
    BuildGUID = BuildGUID & AddZero(Hex$(.Data2), 4) & "-"
    BuildGUID = BuildGUID & AddZero(Hex$(.Data3), 4) & "-"
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

Private Function QWORD(ByVal lngHiDWORD As Long, ByVal lngLoDWORD As Long) As Variant
'By Sebastian Mares, sebastian@maresweb.net, 20040527
Dim curQWORD As Currency
CopyMemory curQWORD, lngLoDWORD, 4
CopyMemory ByVal VarPtr(curQWORD) + 4, lngHiDWORD, 4
QWORD = CDbl(curQWORD) * 10000
End Function

Private Function ExtractDate(Intervals As Variant) As Date
'The value is given as the number of 100-nanosecond intervals since January 1, 1601,
'according to Coordinated Universal Time (Greenwich Mean Time).
Dim varSeconds As Variant
Dim varDays As Variant
varSeconds = CVar(Intervals * ((10 ^ -7)))
varDays = varSeconds / 86400
Debug.Print varSeconds, varDays
ExtractDate = DateAdd("d", varDays, "1/1/1601")
Debug.Print "ExtractDate = "; ExtractDate
End Function
