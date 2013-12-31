VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSHA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       Secure Hash Algorithm-1 (SHA-1)
'
' Module:        clsSHA1.cls
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   The Secure Hash Algorithm (SHA) is required for use with
'                the Digital Signature Algorithm (DSA) as specified in the
'                Digital Signature Standard (DSS) and whenever a secure
'                hash algorithm is required for federal applications.  For
'                a message of length < 2^64 bits, this algorithm produces a
'                condensed representation of the message called a
'                message digest. The message digest is used during
'                generation of a signature for the message.  This
'                also used to compute a message digest for the received
'                version of the message during the process of verifying the
'                signature.  Any change to the message in transit will,
'                with very high probability, result in a different message
'                digest, and the signature will fail to verify.
'
'                This algorithm has been tested to be accurate in
'                accordance with FIPS 180-3 publication dated October 2008.
'                Also, test vectors by Aaron Gifford at
'                http://www.adg.us/computers/sha.html
'
' REFERENCE:     NIST (National Institute of Standards and Technology)
'                (FIPS, Special Publications)
'                http://csrc.nist.gov/publications/PubsFIPS.html
'
'                FIPS 180-2 (Federal Information Processing Standards Publication)
'                dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004
'                http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf
'
'                FIPS 180-3 (Federal Information Processing Standards
'                Publication) dated Oct-2008 (supercedes FIPS 180-2)
'                http://csrc.nist.gov/publications/fips/fips180-3/fips180-3_final.pdf
'
'                Examples of SHA-1, SHA-224, SHA-256, SHA-384 and SHA-512 are available at
'                http://csrc.nist.gov/groups/ST/toolkit/examples.html
'
' Warning:       Feb-2005:  SHA-1 has been compromised.  Recommended that
'                you do not use for password or document authentication.
'                http://www.schneier.com/blog/archives/2005/02/sha1_broken.html
'                http://csrc.nist.gov/groups/ST/toolkit/documents/shs/NISTHashComments-final.pdf
'
'                March 15, 2006:  The SHA-2 family of hash functions
'                (i.e., SHA-224, SHA-256, SHA-384 and SHA-512) may be used
'                by Federal agencies for all applications using secure hash
'                algorithms. Federal agencies should stop using SHA-1 for
'                digital signatures, digital time stamping and other
'                applications that require collision resistance as soon as
'                practical, and must use the SHA-2 family of hash functions
'                for these applications after 2010. After 2010, Federal
'                agencies may use SHA-1 only for the following applications:
'                    - hash-based message authentication codes (HMACs)
'                    - key derivation functions (KDFs)
'                    - random number generators (RNGs)
'                Regardless of use, NIST encourages application and protocol
'                designers to use the SHA-2 family of hash functions for all
'                new applications and protocols.
'                http://csrc.nist.gov/groups/ST/toolkit/secure_hashing.html
'
' ===========================================================================
'
'  SHA1 test results using one (1) pass
'
'   "a"                          - 86f7e437faa5a7fce15d1ddcb9eaeaea377667b8
'   "abc"                        - a9993e364706816aba3e25717850c26c9cd0d89d
'   "message digest"             - c12252ceda8be8994d5fa0290a47231c1d16aae3
'   "abcdefghijklmnopqrstuvwxyz" - 32d10c7b8cf96570ca04ce37f2a19d84240d3a89
'   "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
'                                - 84983e441c3bd26ebaae4aa1f95129e5e54670f1
'   "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'                                - 761c457bf73b14d27e9e9265c46f4b4dda11f940
'   8 times "1234567890"         - 50abf5706a150990a08b2c5ea40fa0e585554732
'   1 million letter "a"         - 34aa973cd4c4daa4f61eeb2bdbad27316534016f
'   1 million zeroes             - bef3595266a65a2ff36b700a75e8ed95c68210b6
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
' 24-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              ReWrote this module
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote LoadWordArray() routine
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 21-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Update reference to FIPS 180-3 and new test vectors
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for notifying me that he had a
'              problem performing a hash with a specific file.  The hashed
'              results were not the same as with other hashing utilities.
'              I found that when a hex value is calculated from a number to
'              hex, a series of letters are created (hex).
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'              The problem arose when converting the hex data back to a long
'              integer.
'                    Ex:  &hA780 -> -22656  (Bad)
'              By appending a trailing ampersand to the hex value the
'              conversion is corrected.
'                    Ex:  &hA780& -> 42880  (Good)
'              See Transform() routine.
'              Added property HashRounds() to set number of rounds for hashing.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while hashing data
'              (Transform() routine).
'              Removed obsolete code.  Updated documentation.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
Private Const MODULE_NAME As String = "cSHA1"
Private Const ZEROES_8    As String = "00000000"

' ***************************************************************************
' Variables
' ***************************************************************************
Private mstrBitsInHex     As String
Private mcurMaxSize       As Currency
Private mcurAccrued       As Currency
Private mcurBitCount      As Currency
Private mcurCurrBitCnt    As Currency
Private mlngRounds        As Long
Private malngConst(80)    As Long
Private malngHash(5)      As Long
Private mobjMath32        As cMath32

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)

Private Sub Class_Initialize()

    ResetVariables
    Set mobjMath32 = Nothing
    ' Free class object from memory
    Set mobjMath32 = New cMath32
    ' Instantiate class object
    HashRounds = 1
    ' Set default number of passes
End Sub

Private Sub Class_Terminate()

    ResetVariables
    Set mobjMath32 = Nothing
End Sub

' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************
' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate the hashed values into one string.
'                SHA-1 output length = 40 bytes (hex format)
'
' Returns:       Concatenated hashed data in byte array.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-SEP-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim intIndex    As Integer
    Dim intPointer  As Integer
    Dim lngProgress As Long
    Dim strHex      As String

    strHex = Space$(64)
    ' preload output buffer
    intPointer = 1

    ' Start position in output buffer
    ' Concatenate the hashed values,
    ' in hex format, into one string
    For intIndex = 0 To 4
        Mid$(strHex, intPointer, 8) = Right$(ZEROES_8 & Hex$(malngHash(intIndex)), 8)
        intPointer = intPointer + 8
    Next
    strHex = StrConv(Trim$(strHex), vbLowerCase)
    ' Convert to lowercase before returning
    FormatOutput = StrConv(strHex, vbFromUnicode)
    ' Convert string data to byte array
    ' 29-Jan-2010 Update progress bar.
    lngProgress = CalcProgress(100, 100)
    RaiseEvent HashProgress(lngProgress)
End Function

' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************
' ***************************************************************************
' Routine:       HashFile
'
' Description:   Function to create a unique hex string representation of 
'                the data passed.  See Readme.txt file for more information. 
' 
' Syntax:
'    With mobjSHA
'        ' convert file location to byte array 
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)
'        abytHashed() = .HashFile(abytData())            ' hash data and return as Byte array
'        txtData.Text = StrConv(abytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' ***************************************************************************
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    On Error GoTo HashFile_CleanUp

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        Exit Function
    End If

    mcurAccrued = 0@
    ' Init accrued progress
    curFilePos = 0@
    ' set to first position in the file
    Erase abytData()
    ' make sure array is empty
    Set objBigFiles = New cBigFiles
    ' Instantiate class object
    ResetVariables
    ' Reset class variables
    Initialize
    ' load the work and constant arrays
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles
        strSource = StrConv(abytFile(), vbUnicode)
        ' convert byte array to string
        .CalcFileSize strSource, curAmtLeft, mstrBitsInHex
        ' Get the size of the file
        mcurMaxSize = curAmtLeft * 2
        ' Capture file size (double for hex chars)
        mcurBitCount = curAmtLeft * 8

        ' Get total bit count for comparison
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        Do
            Erase abytData()
            ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)
            ' Process the source file in chunks
            ReDim abytData(lngBlockSize - 1) As Byte
            ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)

            ' Adjust amount of file left to process
            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)

            ' Process the message data into the appropriate
            ' string length and perform the bit manipulations.
            If Not LoadWordArray(abytData()) Then
                gblnStopProcessing = True
            End If

            ' See if user opted to STOP processing
            DoEvents

            If gblnStopProcessing Then
                Exit Do
            End If

            ' If nothing else to read then leave
            DoEvents

            If curAmtLeft < 1 Then
                Exit Do
            End If

        Loop
        .API_CloseFile hFile
        ' Close all files opened by this routine
    End With

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' Concatenate the state array into one data
    ' string and return the digital signature.
    HashFile = FormatOutput()
HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile
    ' Close all files opened by this routine
    ResetVariables
    ' Reset class variables
    Erase abytData()
    ' Always empty arrays when not needed
    Set objBigFiles = Nothing
    ' Always free objects from memory
End Function

' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************
Public Property Let HashRounds(ByVal lngData As Long)

    ' I have made allowances for up to ten (10) passes.
    Select Case lngData

        Case 1 To 10
            ' Good choice
            mlngRounds = lngData

        Case Else
            ' Use default value
            mlngRounds = 1
    End Select
End Property

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to create a unique hex string representation of 
'                the data passed.  See Readme.txt file for more information. 
' 
' Syntax:
'    With mobjSHA
'        abytData() = StrConv("abc", vbFromUnicode)      ' convert string data to byte array 
'        abytHashed() = .HashString(abytData())          ' hash data and return as Byte array
'        txtData.Text = StrConv(abytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - data string in byte array to be hashed 
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' ***************************************************************************
Public Function HashString(ByRef abytData() As Byte) As Byte()

    On Error GoTo HashString_CleanUp

    ResetVariables
    ' Reset class variables
    Initialize
    ' load the work and constant arrays
    mcurAccrued = 0@
    ' Init accrued progress
    mcurMaxSize = CCur(UBound(abytData) * 2)
    ' Capture string size (double for hex chars)
    mcurBitCount = CCur((UBound(abytData) + 1) * 8)
    ' Get the bit count for the message
    mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)
    ' Convert bit count into hex
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo HashString_CleanUp
    End If

    ' Process the message data into the appropriate
    ' string length and perform the bit manipulations.
    If LoadWordArray(abytData()) Then
        ' Concatenate the state array into one data
        ' string and return the digital signature.
        HashString = FormatOutput()
    Else
        gblnStopProcessing = True
    End If

HashString_CleanUp:
    ResetVariables
    ' Reset class variables
End Function

' ***************************************************************************
' Routine:       Initialize
'
' Description:   Preloads the arrays and constants with predetermined values
'                to be used in the hashing calculations.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Initialize()

    Dim lngIndex As Long

    Erase malngConst()
    Erase malngHash()
    malngHash(0) = &H67452301
    malngHash(1) = &HEFCDAB89
    malngHash(2) = &H98BADCFE
    malngHash(3) = &H10325476
    malngHash(4) = &HC3D2E1F0

    ' Preload the constants array
    For lngIndex = 0 To 79

        Select Case lngIndex

            Case 0 To 19
                malngConst(lngIndex) = &H5A827999

            Case 20 To 39
                malngConst(lngIndex) = &H6ED9EBA1

            Case 40 To 59
                malngConst(lngIndex) = &H8F1BBCDC

            Case 60 To 79
                malngConst(lngIndex) = &HCA62C1D6
        End Select

    Next
End Sub

' *************************************************************************** 
' Routine:       LoadWordArray 
' 
' Description:   All data is stored in Big_Endian format with the Most 
'                Significant Bit (MSB) first. 
' 
'                The message 'M' shall be padded before hash computation
'                begins. The purpose of this padding is to ensure that the
'                padded message is a multiple of 512 bits.
'
'         32-Bit Format
'                Suppose the length of the message 'M', in bits, is 'l'
'                bits. Append the bit '1' to the end of the message,
'                followed by 'k' zero bits, where 'k' is the smallest
'                non-negative solution to the equation l+1+k=448 mod 512.
'                Then append the 64-bit block that is equal to the number
'                'l' expressed using a binary representation. The length of
'                the padded message should now be a multiple of 512 bits.
' 
' Parameters:    abytData() - message to be hashed in a byte array 
' 
' Returns:       TRUE - process completed successfully.
'                FALSE - user reqested to stop processing or an error
'                      occurred.
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 21-Feb-2008  Kenneth Ives  kenaso@tx.rr.com 
'              Rewrote routine 
' *************************************************************************** 
Private Function LoadWordArray(ByRef abytData() As Byte) As Boolean

    Dim lngIndex       As Long
    Dim lngPadding     As Long
    Dim lngMaxBytes    As Long
    Dim lngPosition    As Long
    Dim lngRecLength   As Long
    Dim lngSpaceAlloc  As Long
    Dim lngMsgLenInHex As Long
    Dim strMsg         As String
    Dim strHex         As String
    Dim strTemp        As String

    On Error GoTo LoadWordArray_Error

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo LoadWordArray_Error
    End If

    lngSpaceAlloc = 512
    lngRecLength = 128
    lngMaxBytes = 64
    LoadWordArray = True
    ' Preset flag to TRUE
    strHex = vbNullString
    strTemp = vbNullString
    lngMsgLenInHex = Len(mstrBitsInHex)
    strMsg = StrConv(abytData(), vbUnicode)
    ' convert to string data
    mcurCurrBitCnt = mcurCurrBitCnt + CCur(Len(strMsg) * 8)

    ' calc the ongoing bit count
    ' Incoming message is in equal multiples
    If Len(strMsg) Mod lngMaxBytes = 0 Then
        ' Process the message in equal lengths
        ' there is no need for any padding at
        ' this time.
        Do
            strTemp = Mid$(strMsg, 1, lngMaxBytes)
            ' Capture chunk of data message
            strMsg = Mid$(strMsg, lngMaxBytes + 1)

            ' Resize data message to what was not used
            ' See if we have some data to work with
            If LenB(strTemp) > 0 Then

                ' the string of data is the exact length needed
                If Len(strTemp) = lngMaxBytes Then
                    strHex = Space$(lngSpaceAlloc)
                    lngPosition = 1

                    ' convert data to hex string
                    For lngIndex = 1 To Len(strTemp)
                        Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                        lngPosition = lngPosition + 2
                    Next
                    strTemp = vbNullString
                    Transform strHex
                    ' See if user opted to STOP processing
                    DoEvents

                    If gblnStopProcessing Then
                        Exit Do
                    End If
                End If

            Else
                strHex = vbNullString
                Exit Do
            End If

        Loop While LenB(strMsg) > 0

        strHex = vbNullString
    End If

    ' if the length of the message is not
    ' in equal multiples
    If LenB(strMsg) > 0 Then
        If Len(strMsg) >= lngMaxBytes Then
            Do
                strTemp = Mid$(strMsg, 1, lngMaxBytes)
                ' Capture chunk of data message
                strMsg = Mid$(strMsg, lngMaxBytes + 1)

                ' Resize data message to what was not used
                If Len(strTemp) = lngMaxBytes Then
                    strHex = Space$(lngSpaceAlloc)
                    lngPosition = 1

                    ' convert data to hex string
                    For lngIndex = 1 To Len(strTemp)
                        Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                        lngPosition = lngPosition + 2
                    Next
                    strTemp = vbNullString
                    Transform strHex
                    ' See if user opted to STOP processing
                    DoEvents

                    If gblnStopProcessing Then
                        Exit Do
                    End If

                Else
                    strHex = vbNullString
                    Exit Do
                End If

            Loop
            strHex = vbNullString
        Else
            strTemp = Mid$(strMsg, 1, lngMaxBytes)
            ' Capture chunk of data message
            strMsg = Mid$(strMsg, lngMaxBytes + 1)
            ' Resize data message to what was not used
        End If

        ' the string of data is less than the length needed
        If LenB(strTemp) > 0 Then
            If LenB(strMsg) = 0 Then
                strTemp = strTemp & Chr$(128)
                strHex = Space$(lngSpaceAlloc)
                lngPosition = 1

                ' convert data to hex string
                For lngIndex = 1 To Len(strTemp)
                    Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                    lngPosition = lngPosition + 2
                Next
                strHex = Trim$(strHex)
                ' remove all leading and trailing blanks
                strTemp = vbNullString

                ' if the len of the hex string plus the element length is less
                ' than or equal to the required record length then append the
                ' appropriate number of zeroes and the hex representation of
                ' the message length
                If Len(strHex) < lngRecLength Then
                    ' Calculate amount of padding needed
                    lngPadding = lngRecLength - Len(strHex)

                    Select Case lngPadding

                        Case lngMsgLenInHex
                            ' If the amount of padding needed is the
                            ' exact length of the bit string then
                            ' append, process and leave
                            strHex = strHex & mstrBitsInHex
                            Transform strHex

                        Case Is < lngMsgLenInHex
                            ' Will need to create two final strings
                            ' String 1
                            strHex = strHex & String$(lngPadding, "0")
                            Transform strHex
                            ' Create a string of zeroes and the
                            ' bit count of the message length.
                            ' String 2
                            strHex = vbNullString
                            lngPadding = lngRecLength - lngMsgLenInHex
                            strHex = String$(lngPadding, "0") & mstrBitsInHex
                            Transform strHex

                        Case Is > lngMsgLenInHex
                            ' Adjust amount of padding to fill in
                            ' the gap between data and number of bits
                            lngPadding = lngRecLength - (Len(strHex) + lngMsgLenInHex)
                            strHex = strHex & String$(lngPadding, "0") & mstrBitsInHex
                            Transform strHex
                    End Select

                    GoTo LoadWordArray_CleanUp
                Else
                    ' Hex string is exact length needed.
                    ' Will need to create an extra string.
                    ' String 1
                    Transform strHex
                    ' Create a string of zeroes and the
                    ' bit count of the message length.
                    ' String 2
                    strHex = vbNullString
                    lngPadding = lngRecLength - lngMsgLenInHex
                    strHex = String$(lngPadding, "0") & mstrBitsInHex
                    Transform strHex
                    GoTo LoadWordArray_CleanUp
                End If
            End If
        End If
    End If

    ' No data found in current work string
    If LenB(strMsg) = 0 Then

        ' See if we have reached the end
        ' of the message file or string
        If mcurCurrBitCnt = mcurBitCount Then
            ' End of message indicator is the starting character followed by a
            ' string of zeroes and the hex representation of the message length
            lngPadding = lngRecLength - (lngMsgLenInHex + 2)
            strHex = Hex$(128) & String$(lngPadding, "0") & mstrBitsInHex
            Transform strHex
        End If
    End If

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        LoadWordArray = False
    End If

LoadWordArray_CleanUp:
    strMsg = vbNullString
    strTemp = vbNullString
    strHex = vbNullString

    On Error GoTo 0

    Exit Function
LoadWordArray_Error:
    ErrorMsg MODULE_NAME, "LoadWordArray", err.Description
    LoadWordArray = False

    Resume LoadWordArray_CleanUp

End Function

Private Sub ResetVariables()

    Erase malngConst()
    ' Always empty arrays when not needed
    Erase malngHash()
    mstrBitsInHex = vbNullString
    mcurCurrBitCnt = 0
    mcurBitCount = 0
End Sub

' ***************************************************************************
' Routine:       Transform
'
' Description:   Perform the SHA-1 main hashing functions here.
'
' Parameters:    strMsg - 128 byte message string to be hashed
' 
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2003  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Passed padded message string for possible speed increase.
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for notifying me that he had a
'              problem performing a hash with a specific file.  The hashed
'              results were not the same as with other hashing utilities.
'              I found that when a hex value is calculated from a number to
'              hex, a series of letters are created (hex).
'
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'
'              The problem arose when converting the hex data back to a long
'              integer.
'
'                    Ex:  &hA780 -> -22656  (Bad)
'
'              By appending a trailing ampersand to the hex value the
'              conversion is corrected.
'
'                    Ex:  &hA780& -> 42880  (Good)
'
'              Karl Peterson comment
'              http://bytes.com/topic/visual-basic-net/answers/387151-ampersands-hexadecimal-numbers-whats-rule
'              "The trailing ampersand tells VB it's a Long value (4 byte),
'              even for values that could be held in an Integer (2 byte).
'              Basically, if it's &h8000-&hFFFF, use the ampersand (actually,
'              it's never a bad idea), assuming you really do want to store
'              it in a 4-byte variable. Otherwise, VB strives to confuse, by
'              twiddling the sign bit for those higher values."
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' ***************************************************************************
Private Sub Transform(ByVal strMsg As String)

    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngPointer  As Long
    Dim lngTemp     As Long
    Dim lngTmp1     As Long
    Dim lngTmp2     As Long
    Dim lngProgress As Long
    Dim HH(5)       As Long
    Dim alngMsg(80) As Long

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo Transform_Cleanup
    End If

    With mobjMath32
        lngPointer = 1
        Erase HH()
        Erase alngMsg()

        ' ******** Message preparation **************************
        ' Load array with 32-bit words (8 chars per element)
        ' 28-Dec-2009  Appended ampersand to hex conversion.
        '              See flowerbox above for more information.
        For lngIndex = 0 To 15
            alngMsg(lngIndex) = Val("&H" & Mid$(strMsg, lngPointer, 8) & "&")
            lngPointer = lngPointer + 8
        Next

        ' These 32-bit words must now be extended through the
        ' initial hashing phase to an additional sixty-four
        ' 32-bit words.
        For lngIndex = 16 To 79
            alngMsg(lngIndex) = .w32Rotate(alngMsg(lngIndex - 3) Xor alngMsg(lngIndex - 8) Xor alngMsg(lngIndex - 14) Xor alngMsg(lngIndex - 16), 1)
        Next
        ' ******** End of message preparation ********************
        ' See if user opted to STOP processing
        DoEvents

        If gblnStopProcessing Then
            GoTo Transform_Cleanup
        End If

        ' Set inital values for the hash operators,
        ' this includes previous hash values.
        For lngIndex = 0 To 4
            HH(lngIndex) = malngHash(lngIndex)
        Next

        ' Perform required number of iterations
        For lngIdx = 1 To mlngRounds

            ' The processing below is as per SHA1's specification.
            For lngIndex = 0 To 79

                Select Case lngIndex

                    Case 0 To 19
                        lngTmp1 = (HH(1) And HH(2)) Or ((Not HH(1)) And HH(3))
                        lngTmp2 = &H5A827999

                    Case 20 To 39
                        lngTmp1 = HH(1) Xor HH(2) Xor HH(3)
                        lngTmp2 = &H6ED9EBA1

                    Case 40 To 59
                        lngTmp1 = (HH(1) And HH(2)) Or (HH(1) And HH(3)) Or (HH(2) And HH(3))
                        lngTmp2 = &H8F1BBCDC

                    Case 60 To 79
                        lngTmp1 = HH(1) Xor HH(2) Xor HH(3)
                        lngTmp2 = &HCA62C1D6
                End Select

                ' For each word we process we run the below hashing function and
                ' set it equal to a, shifting the previous a's value down to b,
                ' so a becomes b, b becomes c, after a Left Rotate of 30,
                ' c becomes d, d becomes e.
                lngTemp = .w32Rotate(HH(0), 5)
                lngTemp = .UnsignedAdd(lngTemp, lngTmp1)
                lngTemp = .UnsignedAdd(lngTemp, HH(4))
                lngTemp = .UnsignedAdd(lngTemp, lngTmp2)
                lngTemp = .UnsignedAdd(lngTemp, alngMsg(lngIndex))
                HH(4) = HH(3)
                HH(3) = HH(2)
                HH(2) = .w32Rotate(HH(1), 30)
                ' rotate left
                HH(1) = HH(0)
                HH(0) = lngTemp
                ' See if user opted to STOP processing
                DoEvents

                If gblnStopProcessing Then
                    Exit For
                End If

            Next
            ' See if user opted to STOP processing
            DoEvents

            If gblnStopProcessing Then
                Exit For
            End If

        Next
        ' See if user opted to STOP processing
        DoEvents

        If gblnStopProcessing Then
            GoTo Transform_Cleanup
        End If

        ' Add manipulated values to their original values
        For lngIndex = 0 To 4
            malngHash(lngIndex) = .UnsignedAdd(malngHash(lngIndex), HH(lngIndex))
        Next
    End With

    ' 29-Jan-2010 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(Int(Len(strMsg) / 4))
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent HashProgress(lngProgress)
Transform_Cleanup:
    Erase HH()
    Erase alngMsg()
End Sub
