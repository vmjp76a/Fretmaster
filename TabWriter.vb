Imports System.IO

''' <summary>
''' TabWriter Class - Generates and manages guitar tablature files
''' Converts selected fretboard notes into standard guitar tab format
''' Allows manual editing and saving of tab files
''' </summary>
Public Class TabWriter

    ''' <summary>
    ''' Represents a single note placement on the fretboard
    ''' </summary>
    Public Class TabNote
        Public Property NoteName As String
        Public Property StringNumber As Integer
        Public Property FretNumber As Integer
        
        Public Sub New(note As String, stringNum As Integer, fretNum As Integer)
            NoteName = note
            StringNumber = stringNum
            FretNumber = fretNum
        End Sub
    End Class

    ''' <summary>
    ''' Creates a tab representation from selected notes
    ''' </summary>
    Public Function GenerateTab(notes As List(Of TabNote), tuning As String()) As String
        Dim tabLines(5) As String
        
        For i = 0 To 5
            tabLines(i) = tuning(i) & "|"
        Next
        
        Dim sortedNotes = notes.OrderBy(Function(n) n.FretNumber).ToList()
        
        Dim maxFret = 24
        For fret = 0 To maxFret
            For stringNum = 0 To 5
                Dim note = sortedNotes.FirstOrDefault(Function(n) n.StringNumber = stringNum + 1 And n.FretNumber = fret)
                If note IsNot Nothing Then
                    tabLines(stringNum) &= fret.ToString()
                Else
                    tabLines(stringNum) &= "-"
                End If
            Next
        Next
        
        Return String.Join(Environment.NewLine, tabLines)
    End Function

    ''' <summary>
    ''' Saves tab content to a text file
    ''' </summary>
    Public Sub SaveTab(tabContent As String, filePath As String)
        Try
            Using writer As New StreamWriter(filePath, False)
                writer.WriteLine("Guitar Tablature")
                writer.WriteLine(tabContent)
            End Using
        Catch ex As Exception
            Throw New Exception("Error saving tab file: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Loads tab content from a text file
    ''' </summary>
    Public Function LoadTab(filePath As String) As String
        Try
            If Not File.Exists(filePath) Then
                Return ""
            End If
            Return File.ReadAllText(filePath)
        Catch ex As Exception
            Throw New Exception("Error loading tab file: " & ex.Message)
        End Try
    End Function

End Class