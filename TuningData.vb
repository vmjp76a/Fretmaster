Module TuningData

    ' This module manages guitar string tunings

    ' Standard tunings
    Public Const StandardEADGBE As String = "E A D G B E"
    Public Const StandardDADGAD As String = "D A D G A D"

    ' Alternate tunings
    Public Const DropD As String = "D A D G B E"
    Public Const OpenG As String = "D G D G B D"
    Public Const OpenD As String = "D A D F# A D"

    ' Custom tunings collection
    Private CustomTunings As New Collection

    ' Function to add a custom tuning
    Public Sub AddCustomTuning(tuningName As String, tuningStrings As String)
        On Error Resume Next
        CustomTunings.Add tuningStrings, tuningName
        On Error GoTo 0
    End Sub

    ' Function to get a custom tuning
    Public Function GetCustomTuning(tuningName As String) As String
        On Error Resume Next
        GetCustomTuning = CustomTunings(tuningName)
        On Error GoTo 0
    End Function

    ' Function to save custom tunings to a file
    Public Sub SaveCustomTunings(filePath As String)
        Dim fileLine As String
        Dim fileNumber As Integer
        fileNumber = FreeFile
        Open filePath For Output As #fileNumber
        For Each tuningName In CustomTunings
            fileLine = tuningName & "," & CustomTunings(tuningName)
            Print #fileNumber, fileLine
        Next tuningName
        Close #fileNumber
    End Sub

    ' Function to load custom tunings from a file
    Public Sub LoadCustomTunings(filePath As String)
        Dim fileLine As String
        Dim fileNumber As Integer
        fileNumber = FreeFile
        Open filePath For Input As #fileNumber
        Do While Not EOF(fileNumber)
            Line Input #fileNumber, fileLine
            Dim parts() As String
            parts = Split(fileLine, ",")
            If UBound(parts) = 1 Then
                AddCustomTuning parts(0), parts(1)
            End If
        Loop
        Close #fileNumber
    End Sub

End Module