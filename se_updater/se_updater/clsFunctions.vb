Imports System.Runtime.InteropServices
Imports System.IO

Public Class clsFunctions
    ' This method accepts two strings that represent two files to 
    ' compare. A return value of 0 indicates that the contents of the files
    ' are the same. A return value of any other value indicates that the 
    ' files are not the same.
    Public Shared Function FileCompare(ByVal file1 As String, ByVal file2 As String) As Boolean
        Dim file1byte As Integer
        Dim file2byte As Integer
        Dim fs1 As FileStream
        Dim fs2 As FileStream

        ' Determine if the same file was referenced two times.
        If (file1 = file2) Then
            ' Return 0 to indicate that the files are the same.
            Return True
        End If

        ' Open the two files.
        fs1 = New FileStream(file1, FileMode.Open, FileAccess.Read)
        fs2 = New FileStream(file2, FileMode.Open, FileAccess.Read)

        ' Check the file sizes. If they are not the same, the files
        ' are not equal.
        If (fs1.Length <> fs2.Length) Then
            ' Close the file
            fs1.Close()
            fs2.Close()

            ' Return a non-zero value to indicate that the files are different.
            Return False
        End If

        ' Read and compare a byte from each file until either a
        ' non-matching set of bytes is found or until the end of
        ' file1 is reached.
        Do
            ' Read one byte from each file.
            file1byte = fs1.ReadByte()
            file2byte = fs2.ReadByte()
        Loop While ((file1byte = file2byte) And (file1byte <> -1))

        ' Close the files.
        fs1.Close()
        fs2.Close()

        ' Return the success of the comparison. "file1byte" is
        ' equal to "file2byte" at this point only if the files are 
        ' the same.
        Return ((file1byte - file2byte) = 0)
    End Function
End Class