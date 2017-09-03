Imports System.Drawing
Module Variables
    Friend startingTime As Date

    Friend fileName As String = "C:\myimage.jpg"
    Friend directory As String = IO.Path.GetDirectoryName(fileName) ' Output directory
    Friend original As Bitmap = Image.FromFile(fileName)
    Friend originalSize As Long = original.Width * original.Height

    Friend i As Integer = 0

    Friend stickList As New HashSet(Of Point)
End Module
