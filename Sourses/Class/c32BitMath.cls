VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMath32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cls32BitMath.cls
'
' Description:   Multiple bit manipulating routines.  Some are not used.  I
'                left the unused ones so I would not have to search for them
'                if I use this module in another application.
'
' ***************************************************************************
'
'  Size definitions used by various disk manufacturers
'
'    Bit        0 or 1
'    Nibble     4 Bits
'    Byte       8 Bits
'    Kibibit    1,024 bits
'    Kilobit    1,000 bits
'    Kibibyte   1,024 bytes
'    Kilobyte   1,000 bytes
'    Mebibit    1,048,576 bits
'    Megabit    1,000,000 bits
'    Mebibyte   1,048,576 bytes
'    Megabyte   1,000,000 bytes
'    Gibibit    1,073,741,824 bits
'    Gigabit    1,000,000,000 bits
'    Gibibyte   1,073,741,824 bytes
'    Gigabyte   1,000,000,000 bytes
'    Tebibit    1,099,511,627,776 bits
'    Terabit    1,000,000,000,000 bits
'    Tebibyte   1,099,511,627,776 bytes
'    Terabyte   1,000,000,000,000 bytes
'    Pebibit    1,125,899,906,842,624 bits
'    Petabit    1,000,000,000,000,000 bits
'    Pebibyte   1,125,899,906,842,624 bytes
'    Petabyte   1,000,000,000,000,000 bytes
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
Private Const MODULE_NAME As String = "c32BitMath"
Private Const MAX_INT     As Long = &H7FFF                            '  32767
Private Const MIN_INT     As Long = &H8000                            ' -32768
Private Const KB_64       As Long = &H10000                           '  65536
Private Const MAX_LONG    As Long = &H7FFFFFFF                        '  2147483647
Private Const MIN_LONG    As Long = &H80000000                        ' -2147483648
Private Const HALF_LONG   As Long = &H3FFFFFFF                        '  1073741823

'  4294967296  (== 4.2 Gig)
Private Const GB_4        As Double = (2 ^ 32)

'  4294967295  (1< 4.2 Gig)
Private Const MAX_DWORD   As Double = (2 ^ 32) - 1
Private Const DBL_HIGH    As Double = (1.999999999998 / MAX_DWORD)    ' 0.000000000465661287415694

' ***************************************************************************
' API Declares
' ***************************************************************************
' The CopyMemory function copies a block of memory from one location to
' another. For overlapped blocks, use the MoveMemory function.
' ***************************************************************************
' Variables
'
' Variable name:     mabytPower2
' Naming standard:   m a byt Power2
'                    - - --- ---------
'                    | |  |    |______ Variable subname
'                    | |  |___________ Data type (Byte)
'                    | |______________ Array designator
'                    |________________ Module level designator
'
' ***************************************************************************
Private mblnLoaded     As Boolean
Private mabytValue(8)  As Byte
Private mabytPower2(8) As Byte

' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************
' ***************************************************************************
' Routine:       BytesToLong
'
' Description:   Convert data from a byte array into a long integer. This
'                routine assumes that the byte array will have at least
'                4 elements.
'
' Parameters:    abytData() - Array to hold the data
'                lngPointer - position to start within the array
'
' Returns:       Long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function BytesToLong(ByRef abytData() As Byte, Optional ByVal lngPointer As Long = 0) As Long
Attribute BytesToLong.VB_Description = "Convert data from a byte array data into a long integer."

    Dim lngValue       As Long

    Const ROUTINE_NAME As String = "BytesToLong"

    On Error GoTo BytesToLong_Error

    ' make sure there is at least one byte
    If UBound(abytData) < 1 Then
        InfoMsg "There is not enough data in the incoming array to " & "convert to a long integer." & vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
        GoTo BytesToLong_CleanUp
    End If

    ' Test pointer value
    Select Case lngPointer

        Case Is < 0
            InfoMsg "The starting position must zero or greater." & vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
            GoTo BytesToLong_CleanUp

        Case Is >= UBound(abytData)
            InfoMsg "Starting position in byte array exceeds size of array." & vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
            GoTo BytesToLong_CleanUp

        Case Is > (UBound(abytData) - 3)
            InfoMsg "Incoming array does not have enough data to convert." & vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
            GoTo BytesToLong_CleanUp
    End Select

    CopyMemory lngValue, abytData(0), 4&
    ' Convert bytes to a long integer
    BytesToLong = lngValue
    ' Return long integer
BytesToLong_CleanUp:

    On Error GoTo 0

    ' Nullify this error trap
    Exit Function
BytesToLong_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, err.Description
    gblnStopProcessing = True

    Resume BytesToLong_CleanUp

End Function

Private Sub Class_Initialize()

    ResetVariables
    StopProcessing = False
End Sub

Private Sub Class_Terminate()

    ResetVariables
End Sub

' ***************************************************************************
' Routine:       LongToHex
'
' Description:   Convert a Long to a Hex string. Handles both positive and
'                negative values.
'
' Parameters:    lngNumber - Number to be converted to string data
'                blnFullWord - Flag to designate if a full word is to be
'                            returned (8 chars) or the adjusted length.
'                            Default is full word (TRUE)
'
' Returns:       Converted number
'                Ex:   2147483647 --> 7fffffff  (2gb)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 02-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic bug concerning length adjustment
' ***************************************************************************
Public Function LongToHex(ByVal lngNumber As Long, Optional ByVal blnFullWord As Boolean = True) As String

    Dim strHex     As String
    Dim lngLength  As Long

    Const ZEROES_8 As String = "00000000"
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        LongToHex = vbNullString
        Exit Function
    End If

    strHex = Hex$(lngNumber)

    ' Convert number to hex string
    If blnFullWord Then
        ' Return all eight characters
        strHex = Right$(ZEROES_8 & strHex, 8)
    Else
        lngLength = Len(strHex)

        ' Get length of hex string
        ' Adjust length to multiples of two
        If (lngLength Mod 2 <> 0) Then
            lngLength = lngLength + 1
        End If

        ' Return adjusted length
        strHex = Right$(ZEROES_8 & strHex, lngLength)
    End If

    LongToHex = StrConv(strHex, vbLowerCase)
    ' Convert to lowercase
End Function

Private Sub ResetVariables()

    Erase mabytPower2()
    ' Always start with empty arrays
    Erase mabytValue()
    mblnLoaded = False
    ' Preset flag to FALSE
End Sub

Public Property Get StopProcessing() As Boolean

    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
End Property

' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************
Public Property Let StopProcessing(ByVal blnData As Boolean)

    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
End Property

' ***************************************************************************
' Routine:       SwapEndianHex
'
' Description:   Converts hex data from Big-Endian to Little-Endian or
'                Little-Endian to Big_Endian format. Used with certain
'                hash algorithms.
'
' Parameters:    strHex - Incoming hex string to be manipulated
'                lngRetLength - Optional - Desired length of the return
'                         string.  Default = 16
'
' Returns:       Hex string in new format
'
' Example:         Big-Endian     to    Little-Endian  to    Big-Endian
'                0123456789ABCDEF --> EFCDAB8967452301 --> 0123456789ABCDEF
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function SwapEndianHex(ByVal strHex As String, Optional ByVal lngRetLength As Long = 16) As String

    Dim strOutput   As String
    Dim intPosition As Integer
    Dim intPointer  As Integer

    ' No negative or zero lengths allowed
    If lngRetLength < 1 Then
        lngRetLength = 2
    End If

    ' Return length must be in multiples of 2
    If lngRetLength Mod 2 <> 0 Then
        lngRetLength = lngRetLength + 1
    End If

    ' Verify the incoming hex string is the
    ' same as the return length including
    ' any leading zeroes.
    strHex = Right$(String$(lngRetLength, "0") & strHex, lngRetLength)
    strOutput = Space$(lngRetLength * 2)
    ' Preload output string
    intPointer = lngRetLength - 1

    ' Initialize pointer
    ' Parse hex string capturing two characters
    ' at a time and inserting them into the
    ' output string
    For intPosition = 1 To lngRetLength Step 2
        Mid$(strOutput, intPosition, 2) = Mid$(strHex, intPointer, 2)
        intPointer = intPointer - 2
    Next
    ' Remove leading and trailing spaces
    ' before returning hex string
    SwapEndianHex = Trim$(strOutput)
End Function

' ***************************************************************************
' Routine:       SwapEndianLong
'
' Description:   Returns a Long with reversed byte order.
'
'                Ex:  SwapEndianLong(&H1245678) -->  2018915346 [= &H78563412]
'
' Parameters:    lngValue - long integer to be converted
'
' Returns:       New long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Sep-2004  Mike Sutton   Mike.Sutton@btclick.com
'              http://www.xbeat.net/vbspeed/c_SwapEndian.htm
' 17-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables
' ***************************************************************************
Public Function SwapEndianLong(ByVal lngValue As Long) As Long
Attribute SwapEndianLong.VB_Description = "Returns a Long with reversed byte order."

    SwapEndianLong = (((lngValue And &HFF000000) \ &H1000000) And &HFF&) Or ((lngValue And &HFF0000) \ &H100&) Or ((lngValue And &HFF00&) * &H100&) Or ((lngValue And &H7F&) * &H1000000)

    If (lngValue And &H80) Then
        SwapEndianLong = SwapEndianLong Or &H80000000
    End If
End Function

' ***************************************************************************
' Routine:       UnsignedAdd
'
' Description:   Function to add two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    dblValue1 - Value of A
'                dblValue2 - Value of B
'
' Returns:       Calculated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names
' ***************************************************************************
Public Function UnsignedAdd(ByVal dblValue1 As Double, ByVal dblValue2 As Double) As Long

    Dim dblTemp As Double

    dblTemp = dblValue1 + dblValue2

    If dblTemp < MIN_LONG Then
        UnsignedAdd = CLng(GB_4 + dblTemp)
    Else

        If dblTemp > MAX_LONG Then
            UnsignedAdd = CLng(dblTemp - GB_4)
        Else
            UnsignedAdd = CLng(dblTemp)
        End If
    End If
End Function

' ***************************************************************************
' Routine:       w32Rotate (Long Integer rotation)
'
' Description:   Shifts the bits to the right or left the specified number
'                of positions and returns the new value.  Bits "falling off"
'                the edge will wrap around on the opposite side.  Some
'                common languages like C/C++ or Java have an operator for
'                this job:  ">>>" or "<<<".
'
' Parameters:    lngValue       - numeric value to be evaluated
'                intBitShift    - Number of positions to be shifted
'                                 Positive value = left rotation
'                                 Negative value = right rotation
'
' Returns:       Reformatted value
'
'                 Number                Binary
' Original:      1234567329   00000111010110111100110100010101
'   Left 5:     -3443500432   11101011011110011010001010100000
'  Right 5:    -14725369324   10101000001110101101111001101000
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan  dermot@bitwisemag.com
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and documented
' ***************************************************************************
Public Function w32Rotate(ByVal lngValue As Long, ByVal intBitShift As Integer) As Long

    Dim intLoop    As Integer
    Dim lngSignBit As Long

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test amount of bit shifting
    Select Case intBitShift

        Case 0

            ' return original value if no shift positions
        Case Is < -31
            lngValue = 0

            ' return zero if too many negative shift positions
        Case Is > 31
            lngValue = 0

            ' return zero if too many positive shift positions
        Case Is > 0

            ' A positive bit shift value means rotate left
            '
            ' Shift left by one bit but take in account of
            ' an overflow error from VB. Mask accordingly.
            For intLoop = 1 To intBitShift
                ' Calculate the sign bit of the result
                lngSignBit = lngValue And &HC0000000
                ' Clear all the Most Significant Bit (MSB), that
                ' would be lost anyway, and also clear the sign bit
                lngValue = lngValue And &H3FFFFFFF
                lngValue = lngValue * 2

                ' rotate MSB into LSB
                If lngSignBit And &H80000000 Then
                    lngValue = lngValue Or 1
                End If

                ' set or clear MSB
                If lngSignBit And &H40000000 Then
                    lngValue = lngValue Or &H80000000
                Else
                    lngValue = lngValue And &H7FFFFFFF
                End If

            Next

        Case Is < 0

            ' A negative bit shift value means rotate right
            '
            ' Shift left by one bit but take in account of
            ' an overflow error from VB. Mask accordingly.
            For intLoop = intBitShift To -1
                ' Calculate the sign bit of the result
                lngSignBit = lngValue And &H80000001
                ' Clear all the Most Significant Bit (MSB), that
                ' would be lost anyway, and also clear the sign bit
                lngValue = lngValue And &H7FFFFFFF
                lngValue = lngValue \ 2

                ' rotate MSB into LSB
                If lngSignBit And 1 Then
                    lngValue = lngValue Or &H80000000
                End If

                ' set or clear the old sign bit
                If lngSignBit And &H80000000 Then
                    lngValue = lngValue Or &H40000000
                Else
                    lngValue = lngValue And &HBFFFFFFF
                End If

            Next
    End Select

    w32Rotate = lngValue
    ' return new value
End Function
