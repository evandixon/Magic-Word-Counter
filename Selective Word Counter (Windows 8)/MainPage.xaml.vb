Imports System.Text.RegularExpressions

' The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

''' <summary>
''' An empty page that can be used on its own or navigated to within a Frame.
''' </summary>
Public NotInheritable Class MainPage
    Inherits Page

    ''' <summary>
    ''' Invoked when this page is about to be displayed in a Frame.
    ''' </summary>
    ''' <param name="e">Event data that describes how this page was reached.  The Parameter
    ''' property is typically used to configure the page.</param>
    Protected Overrides Sub OnNavigatedTo(e As Navigation.NavigationEventArgs)
    
    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        Dim text As String = ""
        Dim l = Windows.ApplicationModel.Store.CurrentApp.LicenseInformation
        rtbMain.Document.GetText(Windows.UI.Text.TextGetOptions.UseCrlf, text)

        Dim isTrial = l.IsTrial

#If DEBUG Then
        isTrial = False
#End If

        If isTrial Then
            Dim c = (New WordCount).GetWordCount(text)
            If c.Value > 50 Then
                lblCount.Text = "Word Count: Over 50 (Purchase app to see actual count)."
                'ColorHighlight()
            Else
                lblCount.Text = "Word Count: " & c.ToString()
                'ColorHighlight()
            End If
        Else
            lblCount.Text = "Word Count: " & (New WordCount).GetWordCount(text).ToString()
            ColorHighlight()
        End If
    End Sub
    Sub ColorHighlight()
        Dim x As RichEditBox = rtbMain
        Dim WordCount As New WordCount

        Dim text As String = ""
        rtbMain.Document.GetText(Windows.UI.Text.TextGetOptions.None, text)

        'text = text.Replace(ChrW(8220), """").Replace(ChrW(8221), """").Replace(vbCrLf, " ") 'Remove smart quotes
        'x.Document.SetText(Windows.UI.Text.TextSetOptions.None, text) 'Ensure document is the same after removing smart quotes

        'Reset colors
        x.Document.Selection.SetRange(0, text.Length - 1)
        x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 0, 0, 0)

        Dim endIndexArticle As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(WordCount.ArticleAdjectiveRegexString, RegexOptions.IgnoreCase)).Matches(text)
            If Not endIndexArticle.ContainsKey(WordCount.SimplifiyString(item.Value)) Then
                x.Document.Selection.SetRange(text.IndexOf(item.Value), text.IndexOf(item.Value) + item.Value.Length)
                endIndexArticle.Add(WordCount.SimplifiyString(item.Value), text.IndexOf(item.Value) + item.Value.Length)
            Else
                x.Document.Selection.SetRange(text.IndexOf(item.Value, endIndexArticle(WordCount.SimplifiyString(item.Value))), text.IndexOf(item.Value, endIndexArticle(WordCount.SimplifiyString(item.Value))) + item.Value.Length)
                endIndexArticle(WordCount.SimplifiyString(item.Value)) = text.IndexOf(item.Value, endIndexArticle(WordCount.SimplifiyString(item.Value))) + item.Value.Length
            End If
            x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 100, 100, 100)
        Next

        Dim endIndexName As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(WordCount.NameRegexString, RegexOptions.None)).Matches(text)
            If True Then 'Not (New Regex(WordCount.NameDisqualifierRegex, RegexOptions.None)).IsMatch(item.Value) Then
                If Not endIndexName.ContainsKey(WordCount.SimplifiyString(item.Value)) Then
                    x.Document.Selection.SetRange(text.IndexOf(item.Value), text.IndexOf(item.Value) + item.Value.Length)
                    endIndexName.Add(WordCount.SimplifiyString(item.Value), text.IndexOf(item.Value) + item.Value.Length)
                Else
                    x.Document.Selection.SetRange(text.IndexOf(item.Value, endIndexName(WordCount.SimplifiyString(item.Value))), text.IndexOf(item.Value, endIndexName(WordCount.SimplifiyString(item.Value))) + item.Value.Length)
                    endIndexName(WordCount.SimplifiyString(item.Value)) = text.IndexOf(item.Value, endIndexName(item.Value)) + item.Value.Length
                End If
                x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 0, 0, 255)
            Else
                'It's not a name.
            End If
        Next

        'Dim endIndexError As New Dictionary(Of String, Integer)
        'For Each item As System.Text.RegularExpressions.Match In (New Regex(WordCount.NameErrorRegexString, RegexOptions.None)).Matches(text)
        '    If Not endIndexError.ContainsKey(WordCount.SimplifiyString(item.Value)) Then
        '        x.Document.Selection.SetRange(text.IndexOf(item.Value), text.IndexOf(item.Value) + item.Value.Length)
        '        endIndexError.Add(WordCount.SimplifiyString(item.Value), text.IndexOf(item.Value) + item.Value.Length)
        '    Else
        '        x.Document.Selection.SetRange(text.IndexOf(item.Value, endIndexError(WordCount.SimplifiyString(item.Value))), text.IndexOf(item.Value, endIndexError(WordCount.SimplifiyString(item.Value))) + item.Value.Length)
        '        endIndexError(WordCount.SimplifiyString(item.Value)) = text.IndexOf(item.Value, endIndexError(item.Value)) + item.Value.Length
        '    End If
        '    x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 255, 0, 255)
        'Next

        Dim endIndexDate As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(WordCount.DateRegexString, RegexOptions.None)).Matches(text)
            If Not endIndexDate.ContainsKey(WordCount.SimplifiyString(item.Value)) Then
                x.Document.Selection.SetRange(text.IndexOf(item.Value), text.IndexOf(item.Value) + item.Value.Length)
                endIndexDate.Add(WordCount.SimplifiyString(item.Value), text.IndexOf(item.Value) + item.Value.Length)
            Else
                x.Document.Selection.SetRange(text.IndexOf(item.Value, endIndexDate(WordCount.SimplifiyString(item.Value))), text.IndexOf(item.Value, endIndexDate(WordCount.SimplifiyString(item.Value))) + item.Value.Length)
                endIndexDate(WordCount.SimplifiyString(item.Value)) = text.IndexOf(item.Value, endIndexDate(item.Value)) + item.Value.Length
            End If
            x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 255, 100, 0) 'DarkOrange?
        Next


        Dim endIndexCitation As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(WordCount.CitationRegexString, RegexOptions.IgnoreCase)).Matches(text)
            If Not endIndexCitation.ContainsKey(WordCount.SimplifiyString(item.Value)) Then
                x.Document.Selection.SetRange(text.IndexOf(item.Value), text.IndexOf(item.Value) + item.Value.Length)
                endIndexCitation.Add(WordCount.SimplifiyString(item.Value), text.IndexOf(item.Value) + item.Value.Length)
            Else
                x.Document.Selection.SetRange(text.IndexOf(item.Value, endIndexCitation(WordCount.SimplifiyString(item.Value))), text.IndexOf(item.Value, endIndexCitation(WordCount.SimplifiyString(item.Value))) + item.Value.Length)
                endIndexCitation(WordCount.SimplifiyString(item.Value)) = text.IndexOf(item.Value, endIndexCitation(WordCount.SimplifiyString(item.Value))) + item.Value.Length
            End If
            x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 100, 100, 100)
        Next

        Dim endIndexQuote As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(WordCount.QuoteRegexString, RegexOptions.IgnoreCase)).Matches(text)
            If Not endIndexQuote.ContainsKey(WordCount.SimplifiyString(item.Value)) Then
                x.Document.Selection.SetRange(text.IndexOf(item.Value), text.IndexOf(item.Value) + item.Value.Length)
                endIndexQuote.Add(WordCount.SimplifiyString(item.Value), text.IndexOf(item.Value) + item.Value.Length)
            Else
                x.Document.Selection.SetRange(text.IndexOf(item.Value, endIndexQuote(WordCount.SimplifiyString(item.Value))), text.IndexOf(item.Value, endIndexQuote(WordCount.SimplifiyString(item.Value))) + item.Value.Length)
                endIndexQuote(WordCount.SimplifiyString(item.Value)) = text.IndexOf(item.Value, endIndexQuote(WordCount.SimplifiyString(item.Value))) + item.Value.Length
            End If
            x.Document.Selection.CharacterFormat.ForegroundColor = Windows.UI.Color.FromArgb(255, 255, 0, 0)
        Next

        rtbMain = x
    End Sub

    Private Sub GoToAbout(sender As Object, e As RoutedEventArgs)
        Window.Current.Content = New About(Me)
    End Sub
End Class
