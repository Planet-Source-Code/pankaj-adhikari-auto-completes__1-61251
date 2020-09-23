Attribute VB_Name = "Module1"
Public Type Mystring
    Myarray As New Collection
End Type
Public Function ToArray(Str As String, Findchar As String) As Mystring
Dim Temp As Mystring
Dim String1 As String
Dim Loop1 As Integer
Dim Isit As Boolean
Dim Sstart As Integer
Dim Slength As Integer
While Len(Str) > 0 And Mid(Str, 1, 1) = ","
If Len(Str) > 0 Then
    If Mid(Str, 1, 1) = "," Then
        Str = Mid(Str, 2, Len(Str) - 1)
    End If
End If
Wend
Sstart = 1
For Loop1 = 1 To Len(Str)
    If Mid(Str, Loop1, 1) = Findchar Then
        Isit = True
        Slength = Loop1 - Sstart
        If Slength = 0 Then Slength = 1
        Temp.Myarray.Add (Mid(Str, Sstart, Slength))
        Sstart = Loop1 + 1
    End If
        
Next Loop1
If Sstart <= Len(Str) Then
            Dim Ds As String
            Ds = Mid(Str, Sstart, Len(Str) - Sstart + 1)
            Temp.Myarray.Add Ds
        End If
ToArray = Temp
End Function
