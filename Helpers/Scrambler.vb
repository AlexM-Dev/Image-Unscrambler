Imports System.Drawing

Module Scrambler
    Public Function Scramble(bmp As Bitmap,
                             stickList As HashSet(Of Point)) As Bitmap
        Dim bmp_new As New Bitmap(bmp)

        Dim rnd As New Random

        For x As Integer = 0 To bmp_new.Width - 1
            For y As Integer = 0 To bmp_new.Height - 1
                If Not stickList.Contains(New Point(x, y)) Then
                    Dim x2 As Integer = rnd.Next(bmp_new.Width)
                    Dim y2 As Integer = rnd.Next(bmp_new.Height)

                    If Not stickList.Contains(New Point(x2, y2)) Then
                        Dim pixel As Color = bmp_new.GetPixel(x, y)
                        Dim pixel2 As Color = bmp_new.GetPixel(x2, y2)

                        bmp_new.SetPixel(x, y, pixel2)
                        bmp_new.SetPixel(x2, y2, pixel)
                    End If
                End If
            Next
        Next

        Return bmp_new
    End Function

    Public Function Compare(img1 As Bitmap, img2 As Bitmap) As HashSet(Of Point)
        Dim result As New HashSet(Of Point)
        If img1.Size = img2.Size Then

            For x As Integer = 0 To img1.Width - 1
                For y As Integer = 0 To img1.Height - 1
                    If img1.GetPixel(x, y) = img2.GetPixel(x, y) Then
                        result.Add(New Point(x, y))
                    End If
                Next
            Next
        End If
        Return result
    End Function
    Friend Sub UpdateConsole()
        While stickList.Count < originalSize
            Console.Title = $"Pixels: {stickList.Count}/{originalSize} | Elapsed: {(Date.Now - startingTime).ToString("\:hh\:mm\:ss")} | Image {i}"

            ' Wait 1 millisecond so we don't freeze the entire program.
            Threading.Thread.Sleep(1)
        End While
    End Sub
End Module
