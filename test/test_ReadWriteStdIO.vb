Imports System
Imports System.Collections.Generic
Imports System.Text


Module TestSTDIO

	Public Sub Main(ByVal args As String())

		'create response as an array so it can be easily converted JSON
		Dim ResponseString() As String = {""}

		' get input from Standard Input
		Dim InputString As String = OpenStandardStreamIn()

		' Do whatever you want with Input, then prepare Response
		'ResponseString(0) = DO_SOMETHING(InputString)
		ResponseString(0) = InputString

		' Send Response to Std Output
		OpenStandardStreamOut(ResponseString)

	End Sub

	Public Function OpenStandardStreamIn() As String
		Dim MsgLength As Integer = 0
		Dim InputData As String = ""
		Dim LenBytes As Byte() = New Byte(3) {} 'first 4 bytes are length
		Dim StdIn As System.IO.Stream = Console.OpenStandardInput() 'open the stream
		StdIn.Read(LenBytes, 0, 4) 'length
		MsgLength = System.BitConverter.ToInt32(LenBytes, 0) 'convert length to Int

		Dim Buffer As Char() = New Char(MsgLength - 1) {} 'create Char array for remaining bytes

		Using Reader As System.IO.StreamReader = New System.IO.StreamReader(StdIn) 'Using to auto dispose of stream reader
			While Reader.Peek() >= 0 'while the next byte is not Null
				Reader.Read(Buffer, 0, Buffer.Length) 'add to the buffer

			End While
		End Using

		InputData = New String(Buffer) 'convert buffer to string

		Return InputData
	End Function

	Private Sub OpenStandardStreamOut(ByVal ResponseData() As String) 'fit the response in an array so it can be JSON'd

		Dim OutputData As String = ""
		Dim LenBytes As Byte() = New Byte(3) {} 'byte array for length
		Dim Buffer As Byte() 'byte array for msg

		' OutputData = Newtonsoft.Json.JsonConvert.SerializeObject(ResponseData) 'convert the array to JSON
		OutputData = ResponseData(0)

		Buffer = System.Text.Encoding.UTF8.GetBytes(OutputData) 'convert the response to byte array
		LenBytes = System.BitConverter.GetBytes(Buffer.Length) 'convert the length of response to byte array

		Using StdOut As System.IO.Stream = Console.OpenStandardOutput() 'Using for easy disposal

			StdOut.WriteByte(LenBytes(0)) 'send the length 1 byte at a time
			StdOut.WriteByte(LenBytes(1))
			StdOut.WriteByte(LenBytes(2))
			StdOut.WriteByte(LenBytes(3))

			For i As Integer = 0 To Buffer.Length - 1 'loop the response out byte at a time
				StdOut.WriteByte(Buffer(i))
			Next

		End Using
	End Sub
End Module
