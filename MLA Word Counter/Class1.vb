Imports System.Text
Imports System.Text.RegularExpressions
Public Class Measurement
    Public Property Value As Integer
    Public Property MarginOfErrorPositive As UInt32
    Public Property MarginOfErrorNegative As UInt32
    Public Sub New(_Value As Integer, _MarginPositive As UInt32, _MarginNegative As UInt32)
        Value = _Value
        MarginOfErrorPositive = _MarginPositive
        MarginOfErrorNegative = _MarginNegative
    End Sub
    Public Overrides Function ToString() As String
        If MarginOfErrorPositive = 0 AndAlso MarginOfErrorNegative = 0 Then
            Return Value.ToString & " words"
        Else
            Return Value.ToString & " words (Margin of Error: " & Value - MarginOfErrorNegative & "-" & Value + MarginOfErrorPositive & ")"
        End If
    End Function
End Class
Public Class WordCount
    'Implements wci1_1.WordCounter
    'Implements wci1_1.VersionClass
    Private Shared Function SplitString(Input As String) As String()
        Return Input.Replace("...", " ").Replace("…", " ").Split(" ")
    End Function
    Public Function NameErrorRegexString() As String
        Return "(\.\s\s|\?(\s{1,2})|\!(\s{1,2}))([A-Z]([a-z]|\-)+)\s" & NameRegexString(False)
    End Function
    Public Function QuoteRegexString() As String
        Return "(\"")((.|\n)*?)(\"")"
    End Function
    Public Function CitationRegexString() As String 'Implements wci1_1.WordCounter.CitationRegexString
        Return ("\(((\""?)([A-Z]|[a-z]|[0-9]|\.|\-|\,|\s)+?(\""?)\s)?(([0-9]+?((\-|\,)(\s?)([0-9]+?))*)|n\.d\.)(\;(\s?)((\""?)([A-Z]|[a-z]|[0-9]|\.|\-|\,|\s)+?(\""?)\s)?(([0-9]+?((\-|\,)(\s?)([0-9]+?))*)|n\.d\.))*\)")
    End Function
    Public Function NamePrefixes() As String
        Return "Mr\.|Ms\.|Mrs\.|Miss|Dr\.|President|Agent|Doctor|Mister|Madame|\%name\:|"
    End Function
    Public Function NameRegexString(Optional IncludeTitles As Boolean = True) As String 'Implements wci1_1.WordCounter.NameRegexString
        Dim n As String = ""
        If IncludeTitles Then n = NamePrefixes()
        Return "(" & n & "[A-Z]([A-Z]|[a-z]|\-)+|[A-Z]\.)((\sof)?\s([A-Z]([A-Z]|[a-z]|\-)+)|\s[A-Z]\.)+"
    End Function
    Public Function PlaceRegexString() As String 'Implements wci1_1.WordCounter.PlaceRegexString
        Return "(Mt.|[A-Z]([A-Z]|[a-z]|\-)+|[A-Z]\.)((\sof)?\s([A-Z]([A-Z]|[a-z]|\-)+)|\s[A-Z]\.)*\,(\s)?(Mt.|[A-Z]([A-Z]|[a-z]|\-)+|[A-Z]\.)((\sof)?\s([A-Z]([A-Z]|[a-z]|\-)+)|\s[A-Z]\.)*?"
    End Function
    Public Function DateRegexString() As String ' Implements wci1_1.WordCounter.DateRegexString
        Return "((January|February|March|April|May|June|July|August|September|October|November|December)(\s)([0-9]{1,2})((th|nd|st)?)(\,)(\s?)([0-9]{2,})|([0-9]{1,2})\s(January|February|March|April|May|June|July|August|September|October|November|December)\s([0-9]+?))"
    End Function
    Public Function ClassPeriodRegexString() As String ' Implements wci1_1.WordCounter.ClassPeriodRegexString
        Return "([A-Z]|[a-z]|[0-9]|\s|-)+?\,\s([0-9])(th|nd|st)\s(Period)"
    End Function
    Public Function ArticleAdjectiveRegexString() As String
        Return "(\s|^)(a|an|the)(\s)"
    End Function
    Private Shared Function RemoveBadCharacters(Input As String) As String
        Dim e As New System.Text.ASCIIEncoding
        Return Input.Replace("…", " ") '.Replace(e.GetChars({97}), """").Replace(e.GetChars({93}), """")
    End Function
    Private Shared Function TrimWord(Word As String)
        Return Word.Trim(".,!?""'-–()".ToCharArray)
    End Function
    Public Function GetWordCount(TextToCount As String) As Measurement ' Implements wci1_1.WordCounter.GetWordCount
        Return GetWordCount(TextToCount, True, True, True, True)
    End Function
    Public Shared Function GetWordCount(TextToCount As String, CheckQuote As Boolean, Optional CheckIni As Boolean = True, Optional CheckCitation As Boolean = True, Optional CheckDate As Boolean = True, Optional CheckNames As Boolean = True, Optional CheckPlace As Boolean = True) As Measurement
        Dim m As New WordCount
        'If GetSetting("mlawc", "keys", "indeft", False) = False Then
        '    Return New Measurement(0, 0, 0)
        '    Exit Function
        'End If
        TextToCount = TextToCount.Replace(ChrW(8220), """").Replace(ChrW(8221), """").Replace(vbCrLf, " ")
        Dim output As New Measurement(0, 0, 0)
        Dim words() As String = SplitString(RemoveBadCharacters(TextToCount))
        For Each word In words
            Dim w As String = TrimWord(word)
            If Not String.IsNullOrEmpty(w) AndAlso Not w.ToLower = "a" AndAlso Not w.ToLower = "an" AndAlso Not w.ToLower = "the" Then
                output.Value += 1
            End If
        Next

        'count quotes as one word
        Dim quotes As New Regex(m.QuoteRegexString, RegexOptions.Compiled And RegexOptions.IgnoreCase)
        'Dim wordsregex As New Regex("\b([a-z]+)", RegexOptions.Compiled And RegexOptions.IgnoreCase)
        Dim wordsInQuote As Integer = 0
        Dim quotesMatches As Integer = 0
        If CheckQuote Then
            quotesMatches = quotes.Matches(RemoveBadCharacters(TextToCount)).Count
            For Each item As Match In quotes.Matches(RemoveBadCharacters(TextToCount))
                'For Each word In SplitString(item.Value)
                '    Dim w As String = TrimWord(word)
                '    If Not String.IsNullOrEmpty(w) AndAlso Not w.ToLower = "a" AndAlso Not w.ToLower = "an" AndAlso Not w.ToLower = "the" Then
                '        wordsInQuote += 1
                '    End If
                'Next
                wordsInQuote += GetWordCount(item.Value, False).Value
            Next
        End If

        'Citations
        Dim citationCount As Integer = 0
        Dim citationWords As Integer = 0
        Dim citation As New System.Text.RegularExpressions.Regex(m.CitationRegexString, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
        If CheckCitation Then
            citationCount = citation.Matches(TextToCount).Count
            For Each item In citation.Matches(TextToCount)
                citationWords += GetWordCount(item.value, True, True, False, True, True).Value
            Next
        End If

        'Check for names
        Dim nameMatches As Integer = 0
        Dim nameWords As Integer = 0
        If CheckNames Then
            Dim nameRegex As New Regex(m.NameRegexString, RegexOptions.Compiled)
            Dim nameError As New Regex(m.NameErrorRegexString, RegexOptions.Compiled)
            output.MarginOfErrorPositive = nameError.Matches(TextToCount).Count
            For Each item In nameRegex.Matches(RemoveBadCharacters(TextToCount))
                nameMatches += 1
                nameWords += GetWordCount(item.value, False, True, False, False, False).Value
            Next
        End If

        'Check for places
        Dim placeMatches As Integer = 0
        Dim placeWords As Integer = 0
        If CheckPlace Then
            Dim placeRegex As New Regex(m.PlaceRegexString, RegexOptions.Compiled)
            For Each item In placeRegex.Matches(RemoveBadCharacters(TextToCount))
                placeMatches += 1
                placeWords += GetWordCount(item.value, False, True, False, False, False, False).Value
            Next
        End If

        'Check for dates
        Dim dateMatches As Integer = 0
        Dim dateWords As Integer = 0
        If CheckDate Then
            Dim dateRegex As New Regex(m.DateRegexString, RegexOptions.Compiled And RegexOptions.IgnoreCase)
            For Each item In dateRegex.Matches(RemoveBadCharacters(TextToCount))
                dateMatches += 1
                dateWords += GetWordCount(item.value, False, False, False, False).Value
            Next
        End If

        'Load extension ini
        If CheckIni AndAlso IO.File.Exists(IO.Path.Combine(Environment.CurrentDirectory, "SpecialRules.ini")) Then
            Dim ini As String() = IO.File.ReadAllLines(IO.Path.Combine(Environment.CurrentDirectory, "SpecialRules.ini"))
            For Each line In ini
                If Not (String.IsNullOrEmpty(line) AndAlso line.StartsWith("[") AndAlso line.EndsWith("]")) AndAlso Not line.StartsWith("#") Then
                    Dim lineRegex As New Regex("(\"")(.+)(\"")(\=)([0-9]+)", RegexOptions.Compiled)
                    Dim term As String = lineRegex.Match(line).Groups(2).Value
                    Dim occurrances As Integer = (New Regex(term, RegexOptions.Compiled).Matches(TextToCount).Count)
                    Dim count As Integer = lineRegex.Match(line).Groups(5).Value
                    output.Value = (output.Value - (GetWordCount(term, True, False).Value * occurrances)) + (count * occurrances)
                End If
            Next
        End If

        'Finalize
        output.Value = output.Value - wordsInQuote + quotesMatches - citationWords - nameWords + nameMatches - dateWords + dateMatches - placeWords + placeMatches
        Return output
    End Function
    Public Function Version() As Version ' Implements wci1_1.WordCounter.Version, wci1_1.VersionClass.Version
        Return New Version(1, 0, 1, 0)
    End Function
End Class
