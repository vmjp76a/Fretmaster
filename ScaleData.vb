Public Class ScaleData
    ' Enum for Scale Types
    Public Enum ScaleType
        Major
        Minor
        MajorPentatonic
        MinorPentatonic
    End Enum

    Private ReadOnly ScaleNotes As Dictionary(Of ScaleType, String()) = New Dictionary(Of ScaleType, String()){
        {ScaleType.Major, New String() {"C", "D", "E", "F", "G", "A", "B"}},
        {ScaleType.Minor, New String() {"C", "D", "E♭", "F", "G", "A♭", "B♭"}},
        {ScaleType.MajorPentatonic, New String() {"C", "D", "E", "G", "A"}},
        {ScaleType.MinorPentatonic, New String() {"C", "E♭", "F", "G", "B♭"}}
    }

    Public Function GetScaleNotes(scaleType As ScaleType) As String()
        Return ScaleNotes(scaleType)
    End Function

    Public Function GetIntervals(scaleType As ScaleType) As Integer()
        Select Case scaleType
            Case ScaleType.Major
                Return New Integer() {2, 2, 1, 2, 2, 2, 1}
            Case ScaleType.Minor
                Return New Integer() {2, 1, 2, 2, 1, 2, 2}
            Case ScaleType.MajorPentatonic
                Return New Integer() {2, 2, 3, 2}
            Case ScaleType.MinorPentatonic
                Return New Integer() {3, 2, 2, 3}
            Case Else
                Return New Integer() {}
        End Select
    End Function

    Public Function GetDegrees(scaleType As ScaleType) As Integer()
        Return Enumerable.Range(1, GetScaleNotes(scaleType).Length).ToArray()
    End Function
End Class