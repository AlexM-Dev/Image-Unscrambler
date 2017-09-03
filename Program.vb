Imports System.Drawing

Module Program
    Sub Main(args() As String)
        startingTime = Date.Now
        Dim thr As New Threading.Thread(AddressOf UpdateConsole)
        thr.Start()

        Do While stickList.Count < originalSize
            Dim scr As Bitmap = Scramble(original, stickList)
            Dim diff As HashSet(Of Point) = Compare(original, scr)
            Dim stickcount_1 As Integer = stickList.Count
            stickList.UnionWith(diff)
            If (stickList.Count - stickcount_1) >= 1 Then
                scr.Save(directory & "\" & i & ".jpg", Imaging.ImageFormat.Jpeg)
                i += 1
            End If
        Loop
    End Sub
End Module