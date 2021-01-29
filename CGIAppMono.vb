Option Explicit On
Imports System
Imports System.IO
Imports System.Text

Module CGIApp
    ' CGI4BAS.Net
    ' CGI Routines For CGI_RemoteUser With .Net
    '     Ported from the original
    ' Jan 2021 - John Dovey <dovey.john@gmail.com>
    '====================================
    ' CGI4VB.BAS
    '====================================
    ' CGI routines used with VB 4.0 (32bit) using STDIN / STDOUT.
    '
    ' Version: 1.4 (December 1996)
    '

    ' Author:  Kevin O'Brien <obrienk@pobox.com>
    '                        <obrienk@ix.netcom.com>

    '
    ' Version 1.4 December 1996
    '   call WriteFile as a Function rather than as a Sub
    '   generate error when url-decoding bad data
    '   remove unused variables
    '
    ' Version 1.3 October 1996
    '   add all standard environment variables per CGI/1.1 specs
    '   pass sEncoded to urlDecode ByVal to preserve encoded data
    '   handle a query string entered with a form, instead of either/or
    '   create separate function for storing name/value pairs
    '
    ' Version 1.2 October 1996
    '   replace HTTP/1.0 headers with Status headers
    '   add sendHeader and sendFooter procedures
    '   decode form name as well as form value
    '   create separate function for url-decoding
    '
    ' Version 1.1 September 1996
    '   add HTTP headers
    '   add SendB procedure to send data without a vbCrLf

    Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
    Declare Function ReadFile Lib "kernel32" _
            (ByVal hFile As Long,
             lpBuffer As String, ' Any,
             ByVal nNumberOfBytesToRead As Long,
             lpNumberOfBytesRead As Long,
             lpOverlapped As String) As Long ' As Any
    Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long,
            ByVal lpBuffer As String,
            ByVal nNumberOfBytesToWrite As Long,
            lpNumberOfBytesWritten As Long,
            lpOverlapped As String) As Long ' As Any
    Declare Function SetFilePointer Lib "kernel32" _
           (ByVal hFile As Long,
           ByVal lDistanceToMove As Long,
           lpDistanceToMoveHigh As Long,
           ByVal dwMoveMethod As Long) As Long
    Declare Function SetEndOfFile Lib "kernel32" _
           (ByVal hFile As Long) As Long

    Public Const STD_INPUT_HANDLE = -10&
    Public Const STD_OUTPUT_HANDLE = -11&
    Public Const FILE_BEGIN = 0&


    ' environment variables
    '
    Public CGI_Accept As String
    Public CGI_AuthType As String
    Public CGI_ContentLength As String
    Public CGI_ContentType As String
    Public CGI_GatewayInterface As String
    Public CGI_PathInfo As String
    Public CGI_PathTranslated As String
    Public CGI_QueryString As String
    Public CGI_Referer As String
    Public CGI_RemoteAddr As String
    Public CGI_RemoteHost As String
    Public CGI_RemoteIdent As String
    Public CGI_RemoteUser As String
    Public CGI_RequestMethod As String
    Public CGI_ScriptName As String
    Public CGI_ServerSoftware As String
    Public CGI_ServerName As String
    Public CGI_ServerPort As String
    Public CGI_ServerProtocol As String
    Public CGI_UserAgent As String

    Public lContentLength As Long   ' CGI_ContentLength converted to Long
    Public hStdIn As Long   ' handle of Standard Input
    Public hStdOut As Long   ' handle of Standard Output
    Public sErrorDesc As String ' constructed error message
    Public sEmail As String ' webmaster's/your email address
    Public sFormData As String ' url-encoded data sent by the server

    Structure Pair
        Dim Name As String
        Dim Value As String
    End Structure

    Public tPair() As Pair           ' array of name=value pairs

    Sub Main()

        On Error GoTo ErrorRoutine
        InitCgi()          ' Load environment vars and perform other initialization
        GetFormData()      ' Read data sent by the server
        CGI_Main()         ' Process and return data to server

EndPgm:
        End ' end program

ErrorRoutine:
        sErrorDesc = Err.Description & " Error Number = " & Str(Err.Number)
        ErrorHandler()
        Resume EndPgm
    End Sub

    Sub ErrorHandler()
        Dim rc As Long
        On Error Resume Next

        ' use SetFilePointer API to reset stdOut to BOF
        ' and SetEndOfFile to reset EOF

        rc = SetFilePointer(hStdOut, 0&, 0&, FILE_BEGIN)
        SendHeader("Internal Error")
        Send("<H1>Error in " & CGI_ScriptName & "</H1>")

        Send("The following internal error has occurred:")
        Send("<PRE>" & sErrorDesc & "</PRE>")
        Send("<I>Please</I> note what you were doing when this problem occurred, ")
        Send("so we can identify and correct it. Write down the Web page you were ")
        Send("using, any data you may have entered into a form or search box, ")
        Send("and anything else that may help us duplicate the problem.")
        Send("Then contact the administrator of this service: ")
        Send("<A HREF=""mailto:" & sEmail & """>")
        Send("<ADDRESS>&lt;" & sEmail & "&gt;</ADDRESS></A>")
        SendFooter()

        rc = SetEndOfFile(hStdOut)

    End Sub

    Sub InitCgi()

        hStdIn = GetStdHandle(STD_INPUT_HANDLE)
        hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)

        ' Test
        'hStdIn = System.Console.OpenStandardInput
        'hStdOut = System.Console.OpenStandardOutput
        ' End Test
        sEmail = "YourEmailAddress@Here"

        '==============================
        ' Get the environment variables
        '==============================
        '
        ' Environment variables will vary depending on the server.
        ' Replace any variables below with the ones used by your server.
        '
        CGI_Accept = Environment.GetEnvironmentVariable("HTTP_ACCEPT")
        CGI_AuthType = Environment.GetEnvironmentVariable("AUTH_TYPE")
        CGI_ContentLength = Environment.GetEnvironmentVariable("CONTENT_LENGTH")
        CGI_ContentType = Environment.GetEnvironmentVariable("CONTENT_TYPE")
        CGI_GatewayInterface = Environment.GetEnvironmentVariable("GATEWAY_INTERFACE")
        CGI_PathInfo = Environment.GetEnvironmentVariable("PATH_INFO")
        CGI_PathTranslated = Environment.GetEnvironmentVariable("PATH_TRANSLATED")
        CGI_QueryString = Environment.GetEnvironmentVariable("QUERY_STRING")
        CGI_Referer = Environment.GetEnvironmentVariable("HTTP_REFERER")
        CGI_RemoteAddr = Environment.GetEnvironmentVariable("REMOTE_ADDR")
        CGI_RemoteHost = Environment.GetEnvironmentVariable("REMOTE_HOST")
        CGI_RemoteIdent = Environment.GetEnvironmentVariable("REMOTE_IDENT")
        CGI_RemoteUser = Environment.GetEnvironmentVariable("REMOTE_USER")
        CGI_RequestMethod = Environment.GetEnvironmentVariable("REQUEST_METHOD")
        CGI_ScriptName = Environment.GetEnvironmentVariable("SCRIPT_NAME")
        CGI_ServerSoftware = Environment.GetEnvironmentVariable("SERVER_SOFTWARE")
        CGI_ServerName = Environment.GetEnvironmentVariable("SERVER_NAME")
        CGI_ServerPort = Environment.GetEnvironmentVariable("SERVER_PORT")
        CGI_ServerProtocol = Environment.GetEnvironmentVariable("SERVER_PROTOCOL")
        CGI_UserAgent = Environment.GetEnvironmentVariable("HTTP_USER_AGENT")

        lContentLength = Val(CGI_ContentLength)   'convert to long
        ReDim tPair(0)                            'initialize name/value array

    End Sub

    Sub GetFormData()
        '====================================================
        ' Get the CGI data from STDIN and/or from QueryString
        ' Store name/value pairs
        '====================================================
        Dim sBuff As String    ' buffer to receive POST method data
        Dim lBytesRead As Long      ' actual bytes read by ReadFile()
        Dim rc As Long      ' return code

        ' Method POST - get CGI data from STDIN
        ' Method GET  - get CGI data from QueryString environment variable
        '
        If CGI_RequestMethod = "POST" Then
            ''''sBuff = String(lContentLength, Chr(0))
            sBuff = Space(lContentLength)
            rc = ReadFile(hStdIn, sBuff, lContentLength, lBytesRead, 0&)
            sFormData = Left(sBuff, lBytesRead)

            ' Make sure posted data is url-encoded
            ' Multipart content types, for example, are not necessarily encoded.
            '
            If InStr(1, CGI_ContentType, "www-form-urlencoded", 1) Then
                StorePairs(sFormData)
            End If
        End If
        StorePairs(CGI_QueryString)
    End Sub

    Sub StorePairs(sData As String)
        '=====================================================================
        ' Parse and decode form data and/or query string
        ' Data is received from server as "name=value&name=value&...name=value"
        ' Names and values are URL-encoded
        '
        ' Store name/value pairs in array tPair(), and decode them
        '
        ' Note: if an element in the query string does not contain an "=",
        '       then it will not be stored.
        '
        ' /cgi-bin/pgm.exe?parm=1   "1" gets stored and can be
        '                               retrieved with getCgiValue("parm")
        ' /cgi-bin/pgm.exe?1        "1" does not get stored, but can be
        '                               retrieved with urlDecode(CGI_QueryString)
        '
        '======================================================================
        Dim pointer As Long      ' sData position pointer
        Dim n As Long      ' name/value pair counter
        Dim delim1 As Long      ' position of "="
        Dim delim2 As Long      ' position of "&"
        Dim lastPair As Long      ' size of tPair() array
        Dim lPairs As Long      ' number of name=value pairs in sData

        lastPair = UBound(tPair)    ' current size of tPair()
        pointer = 1
        Do
            delim1 = InStr(CInt(pointer), sData, "=")
            If delim1 = 0 Then Exit Do
            pointer = delim1 + 1
            lPairs += 1
        Loop

        If lPairs = 0 Then Exit Sub  'nothing to add

        ' redim tPair() based on the number of pairs found in sData
        ReDim Preserve tPair(lastPair + lPairs)

        ' assign values to tPair().name and tPair().value
        pointer = 1
        For n = (lastPair + 1) To UBound(tPair)
            delim1 = InStr(CInt(pointer), sData, "=") ' find next equal sign
            If delim1 = 0 Then Exit For         ' parse complete

            tPair(n).Name = UrlDecode(Mid$(sData, pointer, delim1 - pointer))

            delim2 = InStr(CInt(delim1), sData, "&")

            ' if no trailing ampersand, we are at the end of data
            If delim2 = 0 Then delim2 = Len(sData) + 1

            ' value is between the "=" and the "&"
            tPair(n).Value = UrlDecode(Mid$(sData, delim1 + 1, delim2 - delim1 - 1))
            pointer = delim2 + 1
        Next n
    End Sub

    Public Function UrlDecode(ByVal sEncoded As String) As String
        '========================================================
        ' Accept url-encoded string
        ' Return decoded string
        '========================================================

        Dim pointer As Long      ' sEncoded position pointer
        Dim pos As Long      ' position of InStr target

        If sEncoded = "" Then Exit Function

        ' convert "+" to space
        pointer = 1
        Do
            pos = InStr(CInt(pointer), sEncoded, "+")
            If pos = 0 Then Exit Do
            Mid$(sEncoded, pos, 1) = " "
            pointer = pos + 1
        Loop

        ' convert "%xx" to character
        pointer = 1

        On Error GoTo errorUrlDecode

        Do
            pos = InStr(CInt(pointer), sEncoded, "%")
            If pos = 0 Then Exit Do

            Mid$(sEncoded, pos, 1) = Chr("&H" & (Mid$(sEncoded, pos + 1, 2)))
            sEncoded = Left$(sEncoded, pos) _
             & Mid$(sEncoded, pos + 3)
            pointer = pos + 1
        Loop
        On Error GoTo 0     'reset error handling
        UrlDecode = sEncoded
        Exit Function

errorUrlDecode:
        '--------------------------------------------------------------------
        ' If this function was mistakenly called with the following:
        '    UrlDecode("100% natural")
        ' a type mismatch error would be raised when trying to convert
        ' the 2 characters after "%" from hex to character.
        ' Instead, a more descriptive error message will be generated.
        '--------------------------------------------------------------------
        If Err.Number = 13 Then      'Type Mismatch error
            Err.Clear()
            Err.Raise(65001, , "Invalid data passed to UrlDecode() function.")
        Else
            Err.Raise(Err.Number)
        End If
        Resume Next
    End Function

    Function GetCgiValue(cgiName As String) As String
        '====================================================================
        ' Accept the name of a pair
        ' Return the value matching the name
        '
        ' tPair(0) is always empty.
        ' An empty string will be returned
        '    if cgiName is not defined in the form (programmer error)
        '    or, a select type form item was used, but no item was selected.
        '
        ' Multiple values, separated by a semi-colon, will be returned
        '     if the form item uses the "multiple" option
        '     and, more than one selection was chosen.
        '     The calling procedure must parse this string as needed.
        '====================================================================
        Dim n As Integer

        GetCgiValue = ""
        For n = 1 To UBound(tPair)
            If UCase(cgiName) = UCase(tPair(n).Name) Then
                If GetCgiValue = "" Then
                    GetCgiValue = tPair(n).Value
                Else             ' allow for multiple selections
                    GetCgiValue = GetCgiValue & ";" & tPair(n).Value
                End If
            End If
        Next n
    End Function

    Sub SendHeader(sTitle As String)
        Send("Status: 200 OK")
        Send("Content-type: text/html" & vbCrLf)
        Send("<HTML><HEAD><TITLE>" & sTitle & "</TITLE></HEAD>")
    End Sub

    Sub SendFooter()
        '==================================
        ' standardized footers can be added
        '==================================
        Send("</BODY></HTML>")
    End Sub

    Sub Send(s As String)
        '======================
        ' Send output to STDOUT
        '======================
        Dim lBytesWritten As Long

        s &= vbCrLf
        WriteFile(hStdOut, s, Len(s), lBytesWritten, 0&)
    End Sub

    Sub SendB(s As String)
        '============================================
        ' Send output to STDOUT without vbCrLf.
        ' Use when sending binary data. For example,
        ' images sent with "Content-type image/jpeg".
        '============================================
        Dim lBytesWritten As Long

        WriteFile(hStdOut, s, Len(s), lBytesWritten, 0&)
    End Sub
    Sub CGI_Main()

        Send("Status: 200 OK")
        Send("Content-type: text/html" & vbCrLf)

        Send("<HTML><HEAD><TITLE>Hello</TITLE></HEAD>")
        Send("<BODY><BASEFONT FACE=""Verdana,Arial,Helvetica"">")
        Send("<H1>Hello!</H1>")
        Send("This may look like an ordinary web page")
        Send("but it is, in fact generated by a CGI")
        Send("script written in Visual Basic!" & vbCrLf)
        Send("And here is the next line of text!")
        Send("</BODY></HTML>")

    End Sub
End Module
