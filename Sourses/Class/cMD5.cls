VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       MD5 (Message-Digest Algorithm)
'
' Module:        clsMD5
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
'        MD5 Message-Digest Algorithm has been compromised
'
'        At the rump session of Crypto 2004 it was announced that Xiaoyun
'        Wang, Dengguo Feng, Xuejia Lai and Hongbo Yu found collisions for
'        MD4, MD5, RIPEMD, and the 128-bit version of HAVAL.
'        http://eprint.iacr.org/2004/199.pdf
'
'        Mar-2005:  Demonstrating a technique for finding MD5 collisions
'        quickly.  Took eight hours on 1.6 GHz computer.
'        http://cryptography.hyperlink.cz/md5/MD5_collisions.pdf
'
'        Jun-2005:  Two researchers from the Institute for Cryptology and
'        IT-Security have generated PostScript files with identical
'        MD5-sums but entirely different (but meaningful!) content.
'        http://www.schneier.com/blog/archives/2005/06/more_md5_collis.html
'
' ***************************************************************************
'
' Description:   This software is provided ``AS IS'' and any express or
'                implied warranties, including, but not limited to, the
'                implied warranties of merchantability and fitness for a
'                particular purpose, are disclaimed.  In no event shall the
'                authors be liable for any direct, indirect, incidental,
'                special, exemplary, or consequential damages (including, but
'                not limited to, procurement of substitute goods or services;
'                loss of use, data, or profits; or business interruption)
'                however caused and on any theory of liability, whether in
'                contract, strict liability, or tort (including negligence
'                or otherwise) arising in any way out of the use of this
'                software, even if advised of the possibility of such damage.
'
'                This is a class which encapsulates a set of MD5 Message
'                Digest functions.  MD5 algorithm produces a 128 bit digital
'                fingerprint (signature) from an dataset of arbitrary length.
'                For details see RFC 1321 (summarized below).  This
'                implementation is derived from the RSA Data Security, Inc.
'                MD5 Message-Digest algorithm reference implementation
'                (originally written in C).
'
'  NOTES:
'      Network Working Group                                    R. Rivest
'      Request for Comments: 1321     MIT Laboratory for Computer Science
'                                             and RSA Data Security, Inc.
'                                                              April 1992
'
'
'                           The MD5 Message-Digest Algorithm
'
'      Summary
'
'         This document describes the MD5 message-digest algorithm. The
'         algorithm takes as input a message of arbitrary length and produces
'         as output a 128-bit "fingerprint" or "message digest" of the input.
'         It is conjectured that it is computationally infeasible to produce
'         two messages having the same message digest, or to produce any
'         message having a given prespecified target message digest. The MD5
'         algorithm is intended for digital signature applications, where a
'         large file must be "compressed" in a secure manner before being
'         encrypted with a private (secret) key under a public-key cryptosystem
'         such as RSA.
'
'         The MD5 algorithm is designed to be quite fast on 32-bit machines. In
'         addition, the MD5 algorithm does not require any large substitution
'         tables; the algorithm can be coded quite compactly.
'
'         The MD5 algorithm is an extension of the MD4 message-digest algorithm
'         1,2]. MD5 is slightly slower than MD4, but is more "conservative" in
'         design. MD5 was designed because it was felt that MD4 was perhaps
'         being adopted for use more quickly than justified by the existing
'         critical review; because MD4 was designed to be exceptionally fast,
'         it is "at the edge" in terms of risking successful cryptanalytic
'         attack. MD5 backs off a bit, giving up a little in speed for a much
'         greater likelihood of ultimate security. It incorporates some
'         suggestions made by various reviewers, and contains additional
'         optimizations. The MD5 algorithm is being placed in the public domain
'         for review and possible adoption as a standard.
'
'         RFC Author:
'         Ronald L.Rivest
'         Massachusetts Institute of Technology
'         Laboratory for Computer Science
'         NE43 -324545    Technology Square
'         Cambridge, MA  02139-1986
'
'      RSA Copyright Notice
'
'         Copyright (C) 1991-2, RSA Data Security, Inc. Created 1991. All
'         rights reserved.
'
'         License to copy and use this software is granted provided that it
'         is identified as the "RSA Data Security, Inc. MD5 Message-Digest
'         Algorithm " in all material mentioning or referencing this software"
'         or this function.
'
'         License is also granted to make and use derivative works provided
'         that such works are identified as "derived from the RSA Data
'         Security, Inc. MD5 Message-Digest Algorithm" in all material
'         mentioning or referencing the derived work.
'
'         RSA Data Security, Inc. makes no representations concerning either
'         the merchantability of this software or the suitability of this
'         software for any particular purpose. It is provided "as is"
'         without express or implied warranty of any kind.
'
'         These notices must be retained in any copies of any part of this
'         documentation and/or software.
'
'  MD5 test results using one (1) pass
'      "a"                          - 0cc175b9c0f1b6a831c399e269772661
'      "abc"                        - 900150983cd24fb0d6963f7d28e17f72
'      "message digest"             - f96b697d7cb7938d525a2f31aaf161d0
'      "abcdefghijklmnopqrstuvwxyz" - c3fcd3d76192e4007dfb496cca67e13b
'      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'                                   - d174ab98d277d9f5a5611c2c9f419d9f
'      8 times "1234567890"         - 57edf4a22be3c955ac49da2e2107b67a
'      1 million letter "a"         - 7707d6ae4e027c70eea2a935c2296f21
'      1 million zeroes             - 879f4bba57ed37c9ec5e5aedf9864698
'
' ***************************************************************************
' Note:   Test vectors obtained from RFC1321.TXT and
'         http://www.nsrl.nist.gov/testdata/
'
'         Most MD5 hashes in Visual Basic will fail to calculate
'         correctly when performing a hash on 1,000,000 letter a's.
'         This one calculates correctly in accordance with results at
'         http://www.nsrl.nist.gov/testdata/ and RFC1321.TXT.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 29-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated flowerbox documentation.
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
Private Const MODULE_NAME As String = "cMD5"

' ***************************************************************************
' Variables
'
' Variable name:     malngHash
' Naming standard:   m a lng Hash
'                    - - --- ---------
'                    | |  |    |______ Variable subname
'                    | |  |___________ Data type (Long)
'                    | |______________ Array designator
'                    |________________ Module level designator
'
' ***************************************************************************
Private mstrBitsInHex     As String
Private mcurMaxSize       As Currency
Private mcurAccrued       As Currency
Private mcurBitCount      As Currency
Private mcurCurrBitCnt    As Currency
Private mlngRounds        As Long
Private malngHash(4)      As Long
Private malngConst()      As Long
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
    ' Set default number of rounds
End Sub

Private Sub Class_Terminate()

    ResetVariables
    Set mobjMath32 = Nothing
End Sub

' ***************************************************************************
' Routines:      Bit manipulating routines for 128 and 256 bit hash strings
'
' Description:   Manipulates the numeric bits
'
' Returns:       Only the value of "a" is changd.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routines
' ***************************************************************************
Private Sub FF(ByRef a As Long, _
               ByVal B As Long, _
               ByVal C As Long, _
               ByVal D As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer, _
               ByVal lngConst As Long)

    Dim lngTemp As Long

    With mobjMath32
        ' #define F(lngWord, y, z) (((lngWord) & (y)) | ((~x) & (z)))
        lngTemp = (B And C) Or ((Not B) And D)
        ' #define FF(a, b, c, d, x, s, ac) { \
        '  (a) += F ((b), (c), (d)) + (x) + (UINT4)(ac); \
        '  (a) = ROTATE_LEFT ((a), (s)); \
        '  (a) += (b); \
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        lngTemp = .UnsignedAdd(lngTemp, lngConst)
        a = .UnsignedAdd(a, lngTemp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, B)
    End With
End Sub

' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************
' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate the hashed values into one string.
'                MD5 output length = 32 bytes (hex format)
'
' Parameters:    malngHash() - Array of hashed data
'
' Returns:       Hex data in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and documented.  Converted data to
'              lowercase prior to leaving routine.
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim strHex      As String
    Dim intIndex    As Integer
    Dim intPointer  As Integer
    Dim lngTemp     As Long
    Dim lngProgress As Long

    strHex = Space$(64)
    ' preload output buffer
    intPointer = 1

    ' Start position in output buffer
    ' Decode 4 x 32-bit words into 16 bytes with
    ' LSB (Least Significant Bit) first each time
    With mobjMath32

        For intIndex = 0 To 3
            ' Convert data back to Big_Endian format
            lngTemp = .SwapEndianLong(malngHash(intIndex))
            Mid$(strHex, intPointer, 8) = .LongToHex(lngTemp)
            intPointer = intPointer + 8
        Next
    End With

    strHex = StrConv(Trim$(strHex), vbLowerCase)
    ' Convert string to lowercase
    FormatOutput = StrConv(strHex, vbFromUnicode)
    ' convert string data to byte array
    ' 29-Jan-2010 Update progress bar.
    lngProgress = CalcProgress(100, 100)
    RaiseEvent HashProgress(lngProgress)
End Function

Private Sub GG(ByRef a As Long, _
               ByVal B As Long, _
               ByVal C As Long, _
               ByVal D As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer, _
               ByVal lngConst As Long)

    Dim lngTemp As Long

    With mobjMath32
        ' #define G(b, c, d) (((b) & (d)) | ((c) & (~d)))
        lngTemp = (B And D) Or (C And (Not D))
        ' #define GG(a, b, c, d, x, s, ac) { \
        '  (a) += F ((b), (c), (d)) + (x) + (UINT4)(ac); \
        '  (a) = ROTATE_LEFT ((a), (s)); \
        '  (a) += (b); \
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        lngTemp = .UnsignedAdd(lngTemp, lngConst)
        a = .UnsignedAdd(a, lngTemp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, B)
    End With
End Sub

' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************
' ***************************************************************************
' Routine:       HashFile
'
' Description:   Function to quickly convert a file into a hex string
'
' Syntax:
'    With mobjMD5
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
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
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
        ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft, mstrBitsInHex
        ' Get the size of the file
        mcurMaxSize = curAmtLeft
        ' Capture file size
        mcurBitCount = curAmtLeft * 8
        ' Get total bit count for comparison
        mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)
        ' Convert bit count into hex
        mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)

        ' Reverse hex values in string
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
    Set objBigFiles = Nothing
    ' Always free objects from memory
    Erase abytData()
    ' Always empty arrays when not needed
    ResetVariables
    ' Reset class variables
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
' Description:   Function to digest a text string and output the result as
'                a string of hexadecimal characters.  Got basic idea from
'                David Ireland of DI Management Services Pty Ltd
'                <www.di-mgt.com.au>.
'
' Syntax:
'    With mobjMD5
'        abytData() = StrConv("abc", vbFromUnicode)      ' convert string data to byte array 
'        abytHashed() = .HashString(abytData())          ' hash data and return as Byte array
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
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
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
    mcurMaxSize = CCur(UBound(abytData))
    ' Capture string size
    mcurBitCount = CCur((UBound(abytData) + 1) * 8)
    ' Get bit count for for comparison
    mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)
    ' Convert bit count into hex
    mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)
    ' Reverse hex values in string
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
' Routine:       HexStringToByteArray
'
' Description:   The message is converted from a hex string to a byte array.
'
' Parameters:    strMsg - hex string to be converted
'
' Returns:       Message data formatted in an array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function HexStringToByteArray(ByVal strMsg As String) As Byte()

    Dim intLength  As Integer
    Dim intIndex   As Integer
    Dim intPointer As Integer
    Dim abytTemp() As Byte

    Erase abytTemp()
    ' Always start with an empty array
    strMsg = Trim$(strMsg)
    ' remove any leading and trailing blanks
    intPointer = 1
    ' Msg pointer position
    intLength = Len(strMsg) \ 2
    ' Calc length
    ReDim abytTemp(intLength) As Byte

    ' Size return array
    ' conver hex to byte
    For intIndex = 0 To intLength - 1
        abytTemp(intIndex) = CByte("&H" & Mid$(strMsg, intPointer, 2))
        intPointer = intPointer + 2
    Next
    HexStringToByteArray = abytTemp()
    ' Return new byte array
    Erase abytTemp()
    ' Always empty arrays when not needed
End Function

Private Sub HH(ByRef a As Long, _
               ByVal B As Long, _
               ByVal C As Long, _
               ByVal D As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer, _
               ByVal lngConst As Long)

    Dim lngTemp As Long

    With mobjMath32
        ' #define H(b, c, d) ((b) ^ (c) ^ (d))
        lngTemp = (B Xor C Xor D)
        ' #define HH(a, b, c, d, x, s, ac) { \
        '  (a) += F ((b), (c), (d)) + (x) + (UINT4)(ac); \
        '  (a) = ROTATE_LEFT ((a), (s)); \
        '  (a) += (b); \
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        lngTemp = .UnsignedAdd(lngTemp, lngConst)
        a = .UnsignedAdd(a, lngTemp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, B)
    End With
End Sub

Private Sub ii(ByRef a As Long, _
               ByVal B As Long, _
               ByVal C As Long, _
               ByVal D As Long, _
               ByVal lngWord As Long, _
               ByVal intBitShift As Integer, _
               ByVal lngConst As Long)

    Dim lngTemp As Long

    With mobjMath32
        ' #define I(b, c, d) ((c) ^ ((b) | (~d)))
        lngTemp = (C Xor (B Or (Not D)))
        ' #define II(a, b, c, d, x, s, ac) { \
        '  (a) += F ((b), (c), (d)) + (x) + (UINT4)(ac); \
        '  (a) = ROTATE_LEFT ((a), (s)); \
        '  (a) += (b); \
        lngTemp = .UnsignedAdd(lngTemp, lngWord)
        lngTemp = .UnsignedAdd(lngTemp, lngConst)
        a = .UnsignedAdd(a, lngTemp)
        a = .w32Rotate(a, intBitShift)
        a = .UnsignedAdd(a, B)
    End With
End Sub

Private Sub Initialize()

    Dim lngIdx  As Long
    Dim strData As String
    Dim vntData As Variant

    Erase malngHash()
    ' Always start with empty arrays
    Erase malngConst()
    ReDim malngConst(64)
    ' Size arrays
    ' Load initialization values
    malngHash(0) = &H67452301
    malngHash(1) = &HEFCDAB89
    malngHash(2) = &H98BADCFE
    malngHash(3) = &H10325476
    vntData = Empty
    ' Always start with empty variants
    strData = vbNullString
    strData = strData & "d76aa478 e8c7b756 242070db c1bdceee f57c0faf 4787c62a a8304613 fd469501 "
    strData = strData & "698098d8 8b44f7af ffff5bb1 895cd7be 6b901122 fd987193 a679438e 49b40821 "
    strData = strData & "f61e2562 c040b340 265e5a51 e9b6c7aa d62f105d 02441453 d8a1e681 e7d3fbc8 "
    strData = strData & "21e1cde6 c33707d6 f4d50d87 455a14ed a9e3e905 fcefa3f8 676f02d9 8d2a4c8a "
    strData = strData & "fffa3942 8771f681 6d9d6122 fde5380c a4beea44 4bdecfa9 f6bb4b60 bebfbc70 "
    strData = strData & "289b7ec6 eaa127fa d4ef3085 04881d05 d9d4d039 e6db99e5 1fa27cf8 c4ac5665 "
    strData = strData & "f4292244 432aff97 ab9423a7 fc93a039 655b59c3 8f0ccc92 ffeff47d 85845dd1 "
    strData = strData & "6fa87e4f fe2ce6e0 a3014314 4e0811a1 f7537e82 bd3af235 2ad7d2bb eb86d391 "
    vntData = Split(strData, Chr$(32))

    ' Load constant array with long integers
    For lngIdx = 0 To 63
        malngConst(lngIdx) = "&H" & vntData(lngIdx)
    Next
    vntData = Empty
    ' Always empty variants when not needed
    strData = vbNullString
End Sub

' ***************************************************************************
' Routine:       LoadWordArray
'
' Description:   All data is stored in Big_Endian format with the Most 
'                Significant Bit (MSB) first. Message data will be
'                converted to Little_Endian format, with the Least
'                Significant Bit (MSB) first, in the Transform()
'                routine prior to being manipulated.
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
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
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
    Dim abytMsg()      As Byte

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
    Erase abytMsg()
    ' Always start with an empty array
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
                    abytMsg() = HexStringToByteArray(strHex)
                    Transform abytMsg()
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
                    abytMsg() = HexStringToByteArray(strHex)
                    Transform abytMsg()
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
                            abytMsg() = HexStringToByteArray(strHex)
                            Transform abytMsg()

                        Case Is < lngMsgLenInHex
                            ' Will need to create two final strings
                            ' String 1
                            strHex = strHex & String$(lngPadding, "0")
                            abytMsg() = HexStringToByteArray(strHex)
                            Transform abytMsg()
                            ' Create a string of zeroes and the
                            ' bit count of the message length.
                            ' String 2
                            strHex = vbNullString
                            lngPadding = lngRecLength - lngMsgLenInHex
                            strHex = String$(lngPadding, "0") & mstrBitsInHex
                            abytMsg() = HexStringToByteArray(strHex)
                            Transform abytMsg()

                        Case Is > lngMsgLenInHex
                            ' Adjust amount of padding to fill in
                            ' the gap between data and number of bits
                            lngPadding = lngRecLength - (Len(strHex) + lngMsgLenInHex)
                            strHex = strHex & String$(lngPadding, "0") & mstrBitsInHex
                            abytMsg() = HexStringToByteArray(strHex)
                            Transform abytMsg()
                    End Select

                    GoTo LoadWordArray_CleanUp
                Else
                    ' Hex string is exact length needed.
                    ' Will need to create an extra string.
                    ' String 1
                    abytMsg() = HexStringToByteArray(strHex)
                    Transform abytMsg()
                    ' Create a string of zeroes and the
                    ' bit count of the message length.
                    ' String 2
                    strHex = vbNullString
                    lngPadding = lngRecLength - lngMsgLenInHex
                    strHex = String$(lngPadding, "0") & mstrBitsInHex
                    abytMsg() = HexStringToByteArray(strHex)
                    Transform abytMsg()
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
            abytMsg() = HexStringToByteArray(strHex)
            Transform abytMsg()
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
    Erase abytMsg()

    ' Always empty arrays when not needed
    On Error GoTo 0

    ' Nullify error trap in this routine
    Exit Function
LoadWordArray_Error:
    ErrorMsg MODULE_NAME, "LoadWordArray", err.Description
    LoadWordArray = False

    Resume LoadWordArray_CleanUp

End Function

Private Sub ResetVariables()

    Erase malngHash()
    Erase malngConst()
    mcurBitCount = 0
    mcurCurrBitCnt = 0
End Sub

' ***************************************************************************
' Routine:       Transform
'
' Description:   Manipulate the data
'
' Parameters:    abytMsg() - Message chunk in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Changed the input parameter to byte array
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' ***************************************************************************
Private Sub Transform(ByRef abytMsg() As Byte)

    Dim H0          As Long
    Dim H1          As Long
    Dim H2          As Long
    Dim H3          As Long
    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngRound    As Long
    Dim lngProgress As Long
    Dim alngWord()  As Long

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo Transform_Cleanup
    End If

    Erase alngWord()
    ReDim alngWord(16) As Long
    lngIndex = 0
    H0 = malngHash(0)
    ' Initialize temp hash variables
    H1 = malngHash(1)
    H2 = malngHash(2)
    H3 = malngHash(3)

    ' ******** Message preparation **************************
    ' convert data to Little-Endian format
    For lngIdx = 0 To 15
        alngWord(lngIdx) = UnsignedJoin(abytMsg(lngIndex + 3), abytMsg(lngIndex + 2), abytMsg(lngIndex + 1), abytMsg(lngIndex))
        lngIndex = lngIndex + 4
    Next
    ' ******** End of message preparation ********************
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo Transform_Cleanup
    End If

    ' Perform required number of iterations
    For lngRound = 1 To mlngRounds
        ' Round 1
        FF H0, H1, H2, H3, alngWord(0), 7, malngConst(0)
        ' 1
        FF H3, H0, H1, H2, alngWord(1), 12, malngConst(1)
        ' 2
        FF H2, H3, H0, H1, alngWord(2), 17, malngConst(2)
        ' 3
        FF H1, H2, H3, H0, alngWord(3), 22, malngConst(3)
        ' 4
        FF H0, H1, H2, H3, alngWord(4), 7, malngConst(4)
        ' 5
        FF H3, H0, H1, H2, alngWord(5), 12, malngConst(5)
        ' 6
        FF H2, H3, H0, H1, alngWord(6), 17, malngConst(6)
        ' 7
        FF H1, H2, H3, H0, alngWord(7), 22, malngConst(7)
        ' 8
        FF H0, H1, H2, H3, alngWord(8), 7, malngConst(8)
        ' 9
        FF H3, H0, H1, H2, alngWord(9), 12, malngConst(9)
        ' 10
        FF H2, H3, H0, H1, alngWord(10), 17, malngConst(10)
        ' 11
        FF H1, H2, H3, H0, alngWord(11), 22, malngConst(11)
        ' 12
        FF H0, H1, H2, H3, alngWord(12), 7, malngConst(12)
        ' 13
        FF H3, H0, H1, H2, alngWord(13), 12, malngConst(13)
        ' 14
        FF H2, H3, H0, H1, alngWord(14), 17, malngConst(14)
        ' 15
        FF H1, H2, H3, H0, alngWord(15), 22, malngConst(15)
        ' 16
        ' See if user opted to STOP processing
        DoEvents

        If gblnStopProcessing Then
            Exit For
        End If

        ' Round 2
        GG H0, H1, H2, H3, alngWord(1), 5, malngConst(16)
        ' 17
        GG H3, H0, H1, H2, alngWord(6), 9, malngConst(17)
        ' 18
        GG H2, H3, H0, H1, alngWord(11), 14, malngConst(18)
        ' 19
        GG H1, H2, H3, H0, alngWord(0), 20, malngConst(19)
        ' 20
        GG H0, H1, H2, H3, alngWord(5), 5, malngConst(20)
        ' 21
        GG H3, H0, H1, H2, alngWord(10), 9, malngConst(21)
        ' 22
        GG H2, H3, H0, H1, alngWord(15), 14, malngConst(22)
        ' 23
        GG H1, H2, H3, H0, alngWord(4), 20, malngConst(23)
        ' 24
        GG H0, H1, H2, H3, alngWord(9), 5, malngConst(24)
        ' 25
        GG H3, H0, H1, H2, alngWord(14), 9, malngConst(25)
        ' 26
        GG H2, H3, H0, H1, alngWord(3), 14, malngConst(26)
        ' 27
        GG H1, H2, H3, H0, alngWord(8), 20, malngConst(27)
        ' 28
        GG H0, H1, H2, H3, alngWord(13), 5, malngConst(28)
        ' 29
        GG H3, H0, H1, H2, alngWord(2), 9, malngConst(29)
        ' 30
        GG H2, H3, H0, H1, alngWord(7), 14, malngConst(30)
        ' 31
        GG H1, H2, H3, H0, alngWord(12), 20, malngConst(31)
        ' 32
        ' See if user opted to STOP processing
        DoEvents

        If gblnStopProcessing Then
            Exit For
        End If

        ' Round 3
        HH H0, H1, H2, H3, alngWord(5), 4, malngConst(32)
        ' 33
        HH H3, H0, H1, H2, alngWord(8), 11, malngConst(33)
        ' 34
        HH H2, H3, H0, H1, alngWord(11), 16, malngConst(34)
        ' 35
        HH H1, H2, H3, H0, alngWord(14), 23, malngConst(35)
        ' 36
        HH H0, H1, H2, H3, alngWord(1), 4, malngConst(36)
        ' 37
        HH H3, H0, H1, H2, alngWord(4), 11, malngConst(37)
        ' 38
        HH H2, H3, H0, H1, alngWord(7), 16, malngConst(38)
        ' 39
        HH H1, H2, H3, H0, alngWord(10), 23, malngConst(39)
        ' 40
        HH H0, H1, H2, H3, alngWord(13), 4, malngConst(40)
        ' 41
        HH H3, H0, H1, H2, alngWord(0), 11, malngConst(41)
        ' 42
        HH H2, H3, H0, H1, alngWord(3), 16, malngConst(42)
        ' 43
        HH H1, H2, H3, H0, alngWord(6), 23, malngConst(43)
        ' 44
        HH H0, H1, H2, H3, alngWord(9), 4, malngConst(44)
        ' 45
        HH H3, H0, H1, H2, alngWord(12), 11, malngConst(45)
        ' 46
        HH H2, H3, H0, H1, alngWord(15), 16, malngConst(46)
        ' 47
        HH H1, H2, H3, H0, alngWord(2), 23, malngConst(47)
        ' 48
        ' See if user opted to STOP processing
        DoEvents

        If gblnStopProcessing Then
            Exit For
        End If

        ' Round 4
        ii H0, H1, H2, H3, alngWord(0), 6, malngConst(48)
        ' 49
        ii H3, H0, H1, H2, alngWord(7), 10, malngConst(49)
        ' 50
        ii H2, H3, H0, H1, alngWord(14), 15, malngConst(50)
        ' 51
        ii H1, H2, H3, H0, alngWord(5), 21, malngConst(51)
        ' 52
        ii H0, H1, H2, H3, alngWord(12), 6, malngConst(52)
        ' 53
        ii H3, H0, H1, H2, alngWord(3), 10, malngConst(53)
        ' 54
        ii H2, H3, H0, H1, alngWord(10), 15, malngConst(54)
        ' 55
        ii H1, H2, H3, H0, alngWord(1), 21, malngConst(55)
        ' 56
        ii H0, H1, H2, H3, alngWord(8), 6, malngConst(56)
        ' 57
        ii H3, H0, H1, H2, alngWord(15), 10, malngConst(57)
        ' 58
        ii H2, H3, H0, H1, alngWord(6), 15, malngConst(58)
        ' 59
        ii H1, H2, H3, H0, alngWord(13), 21, malngConst(59)
        ' 60
        ii H0, H1, H2, H3, alngWord(4), 6, malngConst(60)
        ' 61
        ii H3, H0, H1, H2, alngWord(11), 10, malngConst(61)
        ' 62
        ii H2, H3, H0, H1, alngWord(2), 15, malngConst(62)
        ' 63
        ii H1, H2, H3, H0, alngWord(9), 21, malngConst(63)
        ' 64
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
    With mobjMath32
        malngHash(0) = .UnsignedAdd(malngHash(0), H0)
        malngHash(1) = .UnsignedAdd(malngHash(1), H1)
        malngHash(2) = .UnsignedAdd(malngHash(2), H2)
        malngHash(3) = .UnsignedAdd(malngHash(3), H3)
    End With

    ' 29-Jan-2010 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(UBound(abytMsg))
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent HashProgress(lngProgress)
Transform_Cleanup:
    Erase alngWord()
    Erase abytMsg()
End Sub

Private Function UnsignedJoin(ByVal a As Byte, ByVal B As Byte, ByVal C As Byte, ByVal D As Byte) As Long

    Dim abytData(4) As Byte

    Erase abytData()
    ' Always start with empty arrays
    abytData(0) = a
    abytData(1) = B
    abytData(2) = C
    abytData(3) = D

    ' Join 4 x 8-bit bytes into one 32-bit word a.b.c.d
    With mobjMath32
        UnsignedJoin = .SwapEndianLong(.BytesToLong(abytData()))
    End With

    Erase abytData()
    ' Always empty arrays when not needed
End Function
