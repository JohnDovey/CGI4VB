' https://support.microsoft.com/en-us/topic/how-to-write-cgi-applications-in-visual-basic-91fdad0b-f900-1c4a-1b34-1b6bb420d9a5
Option Explicit

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

Public Const CGI_AUTH_TYPE         As String = "AUTH_TYPE"
Public Const CGI_CONTENT_LENGTH    As String = "CONTENT_LENGTH"
Public Const CGI_CONTENT_TYPE      As String = "CONTENT_TYPE"
Public Const CGI_GATEWAY_INTERFACE As String = "GATEWAY_INTERFACE"
Public Const CGI_HTTP_ACCEPT       As String = "HTTP_ACCEPT"
Public Const CGI_HTTP_REFERER      As String = "HTTP_REFERER"
Public Const CGI_HTTP_USER_AGENT   As String = "HTTP_USER_AGENT"
Public Const CGI_PATH_INFO         As String = "PATH_INFO"
Public Const CGI_PATH_TRANSLATED   As String = "PATH_TRANSLATED"
Public Const CGI_QUERY_STRING      As String = "QUERY_STRING"
Public Const CGI_REMOTE_ADDR       As String = "REMOTE_ADDR"
Public Const CGI_REMOTE_HOST       As String = "REMOTE_HOST"
Public Const CGI_REMOTE_USER       As String = "REMOTE_USER"
Public Const CGI_REQUEST_METHOD    As String = "REQUEST_METHOD"
Public Const CGI_SCRIPT_NAME       As String = "SCRIPT_NAME"
Public Const CGI_SERVER_NAME       As String = "SERVER_NAME"
Public Const CGI_SERVER_PORT       As String = "SERVER_PORT"
Public Const CGI_SERVER_PROTOCOL   As String = "SERVER_PROTOCOL"
Public Const CGI_SERVER_SOFTWARE   As String = "SERVER_SOFTWARE"

Public Declare Function Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long) As Long

Public Declare Function stdin Lib "kernel32" Alias "GetStdHandle" _
(Optional ByVal Handletype As Long = STD_INPUT_HANDLE) As Long

Public Declare Function stdout Lib "kernel32" Alias "GetStdHandle" _
(Optional ByVal Handletype As Long = STD_OUTPUT_HANDLE) As Long

Public Declare Function ReadFile Lib "kernel32" _
(ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
lpNumberOfBytesRead As Long, Optional ByVal lpOverlapped As Long = 0&) As Long

Public Declare Function WriteFile Lib "kernel32" _
(ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long

Sub Main()

    Dim sReadBuffer As String
    Dim sWriteBuffer As String
    Dim lBytesRead As Long
    Dim lBytesWritten As Long
    Dim hStdIn As Long
    Dim hStdOut As Long
    Dim iPos As Integer

    ' sleep for one minute so the debugger can attach and set a break
    ' point on line below
    ' Sleep 60000

    sReadBuffer = String$(CLng(Environ$(CGI_CONTENT_LENGTH)), 0)

    ' Get STDIN handle
    hStdIn = stdin()
    ' Read client's input
    ReadFile hStdIn, sReadBuffer, Len(sReadBuffer), lBytesRead

    ' Find '=' in the name/value pair and parse the buffer
    iPos = InStr(sReadBuffer, "=")
    sReadBuffer = Mid$(sReadBuffer, iPos + 1)

    ' Construct and send response to the client
    sWriteBuffer = "HTTP/1.0 200 OK" & vbCrLf & "Content-Type: text/html" & _
                    vbCrLf & vbCrLf & "Hello "
    hStdOut = stdout()
    WriteFile hStdOut, sWriteBuffer, Len(sWriteBuffer) + 1, lBytesWritten
    WriteFile hStdOut, sReadBuffer, Len(sReadBuffer), lBytesWritten

End Sub
