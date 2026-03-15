Imports System

Public Class ChordData

    ' Enum to define chord types
    Public Enum ChordType
        Major
        Minor
        Diminished
        Augmented
        Seventh
    End Enum

    ' Method to generate chord tones
    Public Shared Function GenerateChord(chordType As ChordType, root As String) As List(Of String)
        Dim intervals As List(Of Integer)
        Select Case chordType
            Case ChordType.Major
                intervals = New List(Of Integer) From {0, 4, 7}
            Case ChordType.Minor
                intervals = New List(Of Integer) From {0, 3, 7}
            Case ChordType.Diminished
                intervals = New List(Of Integer) From {0, 3, 6}
            Case ChordType.Augmented
                intervals = New List(Of Integer) From {0, 4, 8}
            Case ChordType.Seventh
                intervals = New List(Of Integer) From {0, 4, 7, 10}
            Case Else
                Return New List(Of String)()
        End Select

        Return CalculateNotes(root, intervals)
    End Function

    ' Method to calculate notes based on root and intervals
    Private Shared Function CalculateNotes(root As String, intervals As List(Of Integer)) As List(Of String)
        Dim noteNames As String() = {"C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"}
        Dim rootIndex As Integer = Array.IndexOf(noteNames, root)
        Dim chordNotes As New List(Of String)()

        If rootIndex < 0 Then Return chordNotes ' Invalid root

        For Each interval In intervals
            Dim noteIndex As Integer = (rootIndex + interval) Mod 12
            chordNotes.Add(noteNames(noteIndex))
        Next
        Return chordNotes
    End Function

End Class