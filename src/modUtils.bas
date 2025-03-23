Attribute VB_Name = "modUtils"
Option Explicit

' Generates a Version 4 UUID (Random UUID)
Function GenerateUUIDv4() As String
    Dim chars As String
    Dim uuid As String
    Dim i As Integer

    chars = "0123456789abcdef"
    uuid = ""

    ' Generate the first 8 characters
    For i = 1 To 8
        uuid = uuid & Mid(chars, Int(Rnd() * 16) + 1, 1)
    Next i
    uuid = uuid & "-"

    ' Generate the next 4 characters
    For i = 1 To 4
        uuid = uuid & Mid(chars, Int(Rnd() * 16) + 1, 1)
    Next i
    uuid = uuid & "-"

    ' Generate the next 4 characters (Version 4 UUID)
    uuid = uuid & "4" ' Set the version to 4
    For i = 1 To 3
        uuid = uuid & Mid(chars, Int(Rnd() * 16) + 1, 1)
    Next i
    uuid = uuid & "-"

    ' Generate the next 4 characters (Variant 8, 9, A, or B)
    uuid = uuid & Mid("89ab", Int(Rnd() * 4) + 1, 1) ' Set the variant
    For i = 1 To 3
        uuid = uuid & Mid(chars, Int(Rnd() * 16) + 1, 1)
    Next i
    uuid = uuid & "-"

    ' Generate the final 12 characters
    For i = 1 To 12
        uuid = uuid & Mid(chars, Int(Rnd() * 16) + 1, 1)
    Next i

    GenerateUUIDv4 = uuid
End Function

