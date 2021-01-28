Option Explicit On
Imports System
Imports System.IO

Module MSTest
    Private Const _cGI_AUTH_TYPE As String = "AUTH_TYPE"
    Private Const _cGI_CONTENT_LENGTH As String = "CONTENT_LENGTH"
    Public Const CGI_CONTENT_TYPE As String = "CONTENT_TYPE"
    Public Const CGI_GATEWAY_INTERFACE As String = "GATEWAY_INTERFACE"
    Public Const CGI_HTTP_ACCEPT As String = "HTTP_ACCEPT"
    Public Const CGI_HTTP_REFERER As String = "HTTP_REFERER"
    Public Const CGI_HTTP_USER_AGENT As String = "HTTP_USER_AGENT"
    Public Const CGI_PATH_INFO As String = "PATH_INFO"
    Public Const CGI_PATH_TRANSLATED As String = "PATH_TRANSLATED"
    Public Const CGI_QUERY_STRING As String = "QUERY_STRING"
    Public Const CGI_REMOTE_ADDR As String = "REMOTE_ADDR"
    Public Const CGI_REMOTE_HOST As String = "REMOTE_HOST"
    Public Const CGI_REMOTE_USER As String = "REMOTE_USER"
    Public Const CGI_REQUEST_METHOD As String = "REQUEST_METHOD"
    Public Const CGI_SCRIPT_NAME As String = "SCRIPT_NAME"
    Public Const CGI_SERVER_NAME As String = "SERVER_NAME"
    Public Const CGI_SERVER_PORT As String = "SERVER_PORT"
    Private Const _cGI_SERVER_PROTOCOL As String = "SERVER_PROTOCOL"
    Public Const CGI_SERVER_SOFTWARE As String = "SERVER_SOFTWARE"

    Public ReadOnly Property CGI_SERVER_PROTOCOL As String
        Get
            Return _cGI_SERVER_PROTOCOL
        End Get
    End Property

    Public ReadOnly Property CGI_AUTH_TYPE As String
        Get
            Return _cGI_AUTH_TYPE
        End Get
    End Property

    Public ReadOnly Property CGI_CONTENT_LENGTH As String
        Get
            Return _cGI_CONTENT_LENGTH
        End Get
    End Property

    Sub Main()

        Dim sReadBuffer As String
        Dim sWriteBuffer As String
        Dim lBytesRead As Long
        Dim lBytesWritten As Long
        Dim hStdIn As Long
        Dim hStdOut As Long
        Dim iPos As Integer

        sReadBuffer = String$(CLng(Environment(CGI_CONTENT_LENGTH)), 0)

        ' Read client's input
        'ReadFile hStdIn, sReadBuffer, Len(sReadBuffer), lBytesRead
        sReadBuffer = System.Console.OpenStandardInput(sReadBuffer.Length)

        ' Find '=' in the name/value pair and parse the buffer
        iPos = InStr(sReadBuffer, "=")
        sReadBuffer = Mid$(sReadBuffer, iPos + 1)

        ' Construct and send response to the client
        sWriteBuffer = "HTTP/1.0 200 OK" & vbCrLf & "Content-Type: text/html" &
                    vbCrLf & vbCrLf & "Hello "


        'WriteFile hStdOut, sWriteBuffer, Len(sWriteBuffer) + 1, lBytesWritten
        'WriteFile hStdOut, sReadBuffer, Len(sReadBuffer), lBytesWritten


        Dim standardOutput = New StreamWriter(Console.OpenStandardOutput())
        standardOutput.AutoFlush = True
        Console.SetOut(standardOutput)
        Console.WriteLine(sWriteBuffer)
        Console.WriteLine(sReadBuffer)

End Module
