VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBigFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
'  Module:     clsBigFiles.cls
'
'  Purpose:    This module calculates the size of a file.  Can handle
'              file sizes greater than 2gb.
'
' Reference:   Richard Newcombe  22-Jan-2007
'              Getting Past the 2 Gb File Limit
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
'
'              How To Seek Past VBA's 2GB File Limit
'              http://support.microsoft.com/kb/189981
'
'              How To Write Data to a File Using WriteFile API
'              http://support.microsoft.com/kb/165942/EN-US/
'
'              File Caching
'              http://msdn.microsoft.com/en-us/library/aa364218(VS.85).aspx
'
'              CreateFile Function
'              http://msdn.microsoft.com/en-us/library/aa363858(VS.85).aspx
'
'              INFO: FILE_FLAG_WRITE_THROUGH and FILE_FLAG_NO_BUFFERING
'              http://support.microsoft.com/kb/99794
'
' Description: The descriptions in this module are excerts from Richard
'              Newcombe's article.
'
'              When working in the IDE, any numbers that are entered are
'              limited to a Long variable type. Actually, as far I've
'              found, the IDE uses Longs for most numeric storage within
'              the projects that you write.
'
'              Okay, so what's the problem with Longs? Well, by definition
'              they are a signed 4-byte variable, in hex &H7FFFFFFF, with a
'              lower limit of -2,147,483,648 and an upper limit of
'              2,147,483,647 (2 Gb). &H80000000 stores the sign of the
'              value. Even when you enter values in Hex, they are stored in
'              a Long.
'
'              Working with random access files, you quite often use a Long
'              to store the filesize and current position, completely
'              unaware that if the file you access is just one byte over
'              the 2 Gb size, you can cause your application to corrupt the
'              file when writing to it.
'
'              Unfortunately, there is no quick fix for this. To get around
'              the problem, you need to write your own file handling
'              module, one that uses windows APIs to open, read, write, and
'              close any file.
'
'              The API's expect the Low and High 32-bit values in unsigned
'              format. Also, the APIs return unsigned values. So, the first
'              thing you have to do is decide on a variable type that you
'              can use to store values higher than 2 Gb. After some serious
'              thought, I decided to use a Currency type (64-bit scaled
'              integer) this gives you a 922,337 gig upper file limit, way
'              bigger that the largest hard drive available today.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
' 03-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Documented and modified
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added flag to OpenReadWrite() operations that will not go
'              through any intermediate cache but will go directly to disk.
'              Thanks to Lorin for pointing this out.
' 15-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() routine with IsPathValid() routine.
' ***************************************************************************
Option Explicit

' ********************************************************************
' Constants
' ********************************************************************
Private Const MODULE_NAME          As String = "cBigFiles"
Private Const FILE_BEGIN           As Long = 0
Private Const FILE_SHARE_READ      As Long = &H1
Private Const GENERIC_READ         As Long = &H80000000
Private Const OPEN_EXISTING        As Long = 3
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_LONG             As Long = &H7FFFFFFF          ' 2,147,483,647
Private Const MAX_DWORD            As Currency = (2 ^ 32) - 1    ' 4,294,967,295

' ***************************************************************************
' API Declares
' ***************************************************************************
' PathFileExists function determines whether a path to a file system
' object such as a file or directory is valid. Returns nonzero if the
' file exists.
' The CreateFile function creates or opens the following objects and
' returns a handle that can be used to access the object:
'
'       Files
'       pipes
'       mailslots
'       communications resources
'       disk devices (Windows NT only)
'       consoles
'       directories (open only)
' CloseHandle invalidates the specified object handle, decrements
' the object’s handle count, and performs object retention checks.
' Once the last handle to an object is closed, the object is removed
' from the operating system. In other words, closes an open file.
' This API is used to set the current position in the open file.
' This function is very important because any reads or writes to
' the file do not automatically forward the file position.
' Use SetEndOfFile to set the new end of the file marker. This
' function moves the end-of-file position for the specified file
' to the current position of the file pointer.
' This function writes data to a file, starting at the position
' indicated by the file pointer. After the write operation has
' been completed, the file pointer is adjusted by the number of
' bytes written.
' This function reads data from a file, starting at the position
' indicated by the file pointer. After the read operation has
' been completed, the file pointer is adjusted by the number of
' bytes read.
' GetFileSize determines the size of the file. The file size is
' given in a 64-bit value that is split into two 32-bit values.
' The high-order half is put into the variable passed as
' lpFileSizeHigh; the low-order half is returned by the function.
' To get the size, you can either put the binary or hexadecimal
' values of the two variables side-by-side, or use the formula
' filesize = lpFileSizeHigh * 2^32 + return value. If an error
' occurs, the function instead returns -1.
' Forces Windows to flush the write cache.
' SetFileAttributes Function sets the attributes for a file or directory.
' If the function succeeds, the return value is nonzero.
Private Declare Function SetFilePointer _
                Lib "kernel32.dll" (ByVal hFile As Long, _
                                    ByVal lDistanceToMove As Long, _
                                    lpDistanceToMoveHigh As Long, _
                                    ByVal dwMoveMethod As Long) As Long

' ***************************************************************************
' Routine:       API_CloseFile
'
' Description:   Closes an open file.
'
' Parameters:    hFile - Numeric value designating an open file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub API_CloseFile(ByRef hFile As Long)

    ' Always close a file when not in use or
    ' undesired consequences may happen.
    If hFile > 0 Then
        CloseHandle hFile
        ' Release file handle
        hFile = 0
        ' Reset handle value
    End If
End Sub

' ***************************************************************************
' Routine:       API_ReadFile
'
' Description:   This routine is used to read data from an opened file.
'
'                This works much like the standard Read # command. You pass
'                the function the file handle you got when opening the file,
'                the byte position you want to start our read from, the size
'                of the data block you require, and a buffer. For the buffer,
'                you use an array of bytes because this seems to be the best
'                method to ensure that any type of data can be read and
'                written with it.
'
'                You use your conversion function to split the position info
'                into the two 32-bit variables needed for the SetFilePointer
'                API. Set the location in your file and then call your
'                ReadFile API. And, for when you read at the end of the file,
'                you pass the number of bytes that have been read.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'                abytData() - Byte array to hold the data that was read
'
' Returns:       TRUE - Successfully read the file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function API_ReadFile(ByVal hFile As Long, ByVal curPosition As Currency, ByRef abytData() As Byte) As Boolean

    Dim lngRetCode     As Long
    Dim lngLowOrder    As Long
    Dim lngHighOrder   As Long
    Dim lngBytesRead   As Long
    Dim lngBytesToRead As Long

    Const ROUTINE_NAME As String = "API_ReadFile"

    On Error GoTo API_ReadFile_Error

    API_ReadFile = False
    ' Preset to bad ending
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        Exit Function
    End If

    ' Get size of data to write
    lngBytesToRead = (UBound(abytData) + 1) * LenB(abytData(0))
    ' Calculate the current position within the file
    Size2Long curPosition, lngLowOrder, lngHighOrder
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo API_ReadFile_CleanUp
    End If

    ' Set the pointer to start at specific position within the file
    SetFilePointer hFile, lngLowOrder, lngHighOrder, FILE_BEGIN
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        GoTo API_ReadFile_CleanUp
    End If

    ' Test for successful file pointer
    If lngHighOrder = &HFFFFFFFF Then
        InfoMsg "Failed to set read file pointer." & vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo API_ReadFile_CleanUp
    End If

    ' Read this portion of the file
    lngRetCode = ReadFile(hFile, abytData(0), lngBytesToRead, lngBytesRead, 0&)

    ' Test for successful file read
    If lngRetCode = 0 Then
        InfoMsg "Failed to read file." & vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo API_ReadFile_CleanUp
    End If

    API_ReadFile = True
    ' Set flag for successful finish
API_ReadFile_CleanUp:
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        API_CloseFile hFile
        ' Verify file handle has been released
        API_ReadFile = False
    End If

    On Error GoTo 0

    ' Nullify this error trap
    Exit Function
API_ReadFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, err.Description
    API_CloseFile hFile
    ' Verify file handle has been released
    API_ReadFile = False

    Resume API_ReadFile_CleanUp

End Function

' ***************************************************************************
' Routine:       CalcFileSize
'
' Description:   This routine is used to open a file as read only and
'                calculate it's size.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName  - Name of file
'                curFileSize  - Returned file size in bytes
'                strBitsInHex - OPTIONAL - Return the file size calculated
'                               into bits and in hex format without leading
'                               zeroes. Used for calculating a hash.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 15-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated file size calculations.
' ***************************************************************************
Public Sub CalcFileSize(ByVal strFileName As String, _
                        ByRef curFilesize As Currency, _
                        Optional ByRef strBitsInHex As String = vbNullString)

    Dim hFile          As Long
    Dim lngLowOrder    As Long
    Dim lngHighOrder   As Long
    Dim curHexTemp     As Currency

    Const ROUTINE_NAME As String = "CalcFileSize"

    ' receives a handle to the file
    ' receive the low-order half of the file size
    ' receive the high-order half of the file size
    ' Holding area if greater than 2gb
    On Error GoTo CalcFileSize_Error

    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        Exit Sub
    End If

    curFilesize = 0@
    lngHighOrder = 0&

    ' Open file as read only and get the file handle
    If OpenReadOnly(strFileName, hFile) Then
        SetFilePointer hFile, 0, 0, FILE_BEGIN
        ' Set pointer to beginning of file
        lngLowOrder = GetFileSize(hFile, lngHighOrder)
        ' Get API file size values
        API_CloseFile hFile
        ' Close file handle
        Long2Size curFilesize, lngLowOrder, lngHighOrder

        ' Calculate actual file size
        '************************
        '* Optional return data *
        '************************
        ' Test for overflow when calculating the
        ' number of bits. This information is
        ' used for generating file hashes.
        If (curFilesize * 8) >= MAX_LONG Then
            ' File size is greater than or equal to 2gb
            curHexTemp = (curFilesize * 8)
            strBitsInHex = NumberToHex(curHexTemp)
        Else
            ' File size is less than 2gb
            strBitsInHex = Hex$(curFilesize * 8)
        End If

    Else
        curFilesize = CCur(INVALID_HANDLE_VALUE)
    End If

CalcFileSize_CleanUp:

    If hFile > 0 Then
        API_CloseFile hFile
    End If

    On Error GoTo 0

    ' Nullify this error trap
    Exit Sub
CalcFileSize_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, err.Description
    curFilesize = 0
    strBitsInHex = vbNullString

    Resume CalcFileSize_CleanUp

End Sub

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Test to see if a folder or file exists.  Determines if a
'                file exists. This function tests the validity of the file
'                and path. This function tests the validity of the path. A
'                path specified by Universal Naming Convention (UNC) is
'                limited to a file only; that is, \\server\share\file is
'                permitted. A UNC path to a server or server share is not
'                permitted; that is, \\server or \\server\share. This
'                function returns FALSE if a mounted remote drive is out of
'                service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function IsPathValid(ByVal strName As String) As Boolean

    IsPathValid = CBool(PathFileExists(strName))
End Function

' ***************************************************************************
' Routine:       Long2Size
'
' Description:   This routine will convert two Long values into one
'                Currency value.  The multiplication here does not
'                give any erroneous results, but note that the currency
'                type variable must be listed first; otherwise, VB will
'                try to use a Long type variable to temporally store
'                the result. (This is one of the minor issues in VB6.
'                When doing calculations, VB uses the same variable
'                type of your first variable in the calculation and not
'                the variable type of the destination variable. This
'                has been documented on several sites.)
'
' Parameters:    curFileSize  - File size to be calculated
'                lngLowOrder  - Highest value = 2,147,483,647
'                lngHighOrder - Highest value = 2,147,630
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Renamed, modified and documented
' ***************************************************************************
Private Sub Long2Size(ByRef curFilesize As Currency, ByVal lngLowOrder As Long, ByVal lngHighOrder As Long)

    curFilesize = MAX_DWORD * CCur(lngHighOrder)

    If lngLowOrder < 0 Then
        curFilesize = curFilesize + (MAX_DWORD + CCur(lngLowOrder + 1))
    Else
        curFilesize = curFilesize + CCur(lngLowOrder)
    End If
End Sub

' ***************************************************************************
' Routine:       NumberToHex
'
' Description:   Convert a whole number to a hex string
'
' Parameters:    vntNumber - Number to be converted
'
' Returns:       Hex string
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function NumberToHex(ByVal vntNumber As Variant) As String

    Dim vntHex As Variant
    Dim strHex As String

    vntHex = Empty
    ' Always start with an empty variant
    strHex = vbNullString
    vntNumber = CDec(vntNumber)

    If vntNumber < 1 Then
        GoTo NumberToHex_CleanUp
    End If

    Do While vntNumber > 0
        vntHex = Fix(vntNumber / 16)
        vntHex = vntNumber - (vntHex * 16)
        vntNumber = Int(vntNumber / 16)

        Select Case vntHex

            Case 10
                strHex = "A" & strHex

            Case 11
                strHex = "B" & strHex

            Case 12
                strHex = "C" & strHex

            Case 13
                strHex = "D" & strHex

            Case 14
                strHex = "E" & strHex

            Case 15
                strHex = "F" & strHex

            Case Else
                strHex = CStr(vntHex) & strHex
        End Select

    Loop
NumberToHex_CleanUp:
    NumberToHex = strHex
    ' Return hex string
    vntHex = Empty
    ' Always empty variants when not needed
End Function

' ***************************************************************************
' ***                           Methods                                   ***
' ***************************************************************************
' ***************************************************************************
' Routine:       OpenReadOnly
'
' Description:   Open a file to be used as input.  The file must already
'                exist.  If the file does not exist, an error will occur.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName - Fully qualified path and file name
'                hFile - Numeric value designating an open file
'
' Returns:       TRUE - Successfully opened file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function OpenReadOnly(ByVal strFileName As String, ByRef hFile As Long) As Boolean

    Const ROUTINE_NAME As String = "OpenReadOnly"

    On Error GoTo OpenReadOnly_Error

    OpenReadOnly = False
    ' Preset to bad ending
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        Exit Function
    End If

    ' make sure the file does exist
    If Not IsPathValid(strFileName) Then
        Exit Function
    End If

    ' Open a file that is to be used as input (read only).
    hFile = CreateFile(strFileName & vbNullChar, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)

    ' error opening the file
    If hFile = INVALID_HANDLE_VALUE Then
        Exit Function
    End If

    ' Set the pointer to start at the beginning of the file
    SetFilePointer hFile, 0, 0, FILE_BEGIN
    ' See if user opted to STOP processing
    DoEvents

    If gblnStopProcessing Then
        API_CloseFile hFile
        ' Verify file handle has been released
    Else
        OpenReadOnly = True
    End If

OpenReadOnly_CleanUp:

    On Error GoTo 0

    ' Nullify this error trap
    Exit Function
OpenReadOnly_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, err.Description
    API_CloseFile hFile
    ' Verify file handle has been released
    OpenReadOnly = False

    Resume OpenReadOnly_CleanUp

End Function

' ***************************************************************************
' ****                Internal Procedures and Functions                  ****
' ***************************************************************************
' ***************************************************************************
' Routine:       Size2Long
'
' Description:   This routine will work out the higher 32 bits. This code
'                looks like it could be done with a simple division, but
'                you have the problem of the IDE using longs. So, rather
'                than running the risk of the IDE using a long somewhere
'                in the calculations (tests have returned varied results
'                where the division is often 1 out on certain file sizes),
'                you may find a division method that will give the correct
'                value every time, but for now this method will suffice.
'
'                With your max filesize being 922,337 Gb, the highest
'                value that LongHigh will hold is 2,147,630, way below the
'                maximum positive value that a Long can hold. This means
'                that you do not need to monitor it.
'
' Parameters:    curFileSize - File size to be evaluated
'                lngLowOrder - The highest value this will hold is 2,147,483,647
'                lngHighOrder - The highest value this will hold is 2,147,630
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub Size2Long(ByVal curFilesize As Currency, ByRef lngLowOrder As Long, ByRef lngHighOrder As Long)

    lngHighOrder = 0
    lngLowOrder = 0

    ' If the filesize is smaller than 4 Gb, the Do loop does not execute.
    ' Between 4 and 8 Gb the Do loop executes once.  Between 8-12 Gb, it
    ' executes twice.
    Do Until curFilesize < MAX_DWORD
        lngHighOrder = lngHighOrder + 1
        curFilesize = curFilesize - MAX_DWORD
    Loop

    ' Take the remainder and decide whether it needs the sign bit of the
    ' long to hold a value (&H80000000). If it's smaller than 2,147,483,648
    ' you can pass the value directly to the low order. If not, you have to
    ' convert the value to the hex equivalent in signed 32bit.
    If curFilesize > MAX_LONG Then
        ' Larger than 2gb
        lngLowOrder = CLng((MAX_DWORD - (curFilesize - 1)) * -1)
    Else
        ' Less than 2gb
        lngLowOrder = CLng(curFilesize)
    End If
End Sub
