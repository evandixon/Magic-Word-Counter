Public Class StatisticsWindow

    Public Overloads Sub Show(TextForStatistics As String)
        Dim words As New Dictionary(Of String, Integer)
        For Each word In TextForStatistics.Trim.Replace(vbCr, " ").Replace(vbLf, " ").Split(" ")
            Dim w As String = word.Trim.Trim("!.,""'?()#*".ToCharArray)
            If Not String.IsNullOrEmpty(w) Then
                If words.ContainsKey(w.ToLower) Then
                    words(w.ToLower) += 1
                Else
                    words.Add(w.ToLower, 1)
                End If
            End If
        Next
        Dim out() As String = {}
        For Each item In words
            ReDim Preserve out(out.Length)
            out(out.Length - 1) = item.Value & " - " & item.Key
        Next
        Array.Sort(out, New Comparer)
        TextBox1.Text = "Word usage:" & vbCrLf
        For Each item In out
            TextBox1.Text += item & vbCrLf
        Next
        'TextBox1.Text += vbCrLf & vbCrLf & "Text for above statistics:" & vbCrLf & TextForStatistics
        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = 0
        MyBase.Show()
    End Sub
End Class
Public Class Comparer
    Implements IComparer(Of String)

    Public Function Compare(x As String, y As String) As Integer Implements System.Collections.Generic.IComparer(Of String).Compare
        Return -(CInt(x.Split(" ")(0)) - CInt(y.Split(" ")(0)))
    End Function
End Class