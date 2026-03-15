Imports System.Media
Imports NAudio.Wave

Public Class SoundPlayer
    Private waveOut As IWaveOut
    Private audioFileReader As AudioFileReader

    Public Sub New()
    End Sub

    Public Sub PlaySound(note As String, stringNumber As Integer, fretNumber As Integer)
        Dim fileName As String = $"C:\Fretmaster\Sounds\{note}{stringNumber}{fretNumber}.wav"
        PlayFile(fileName)
    End Sub

    Private Sub PlayFile(filePath As String)
        If System.IO.File.Exists(filePath) Then
            waveOut = New WaveOutEvent()
            audioFileReader = New AudioFileReader(filePath)
            waveOut.Init(audioFileReader)
            waveOut.Play()
        Else
            Throw New FileNotFoundException($"Audio file '{filePath}' not found.")
        End If
    End Sub

    Public Sub StopSound()
        If waveOut IsNot Nothing Then
            waveOut.Stop()
            waveOut.Dispose()
        End If
    End Sub
End Class

