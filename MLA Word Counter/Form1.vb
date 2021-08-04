Imports System.Text.RegularExpressions
Imports CryptoLibrary
Imports CryptoLibrary.Formats
Imports Magic_Word_Counter.WordCount
Public Class Form1
    Dim ReCountOnSelectionChange As Boolean = True
    Dim ColorHighlightOnTimerTick As Boolean = True
    Dim Counter As WordCount
    Dim oldText As String = ""
    'Function DoUpdate(Optional IgnoreCurrentVersion As Boolean = False, Optional ShowDialogIfNoUpdate As Boolean = False) As Boolean
    '    Dim output As Boolean = True
    '    Try
    '        Dim us As New MLAWCUpdater.MLAWCUpdaterSoapClient
    '        SaveSetting("mlawc", "keys", "indeft", us.IndefinateTrialValid)
    '        Dim IsLatestVersion As Boolean
    '        If IgnoreCurrentVersion Then
    '            IsLatestVersion = False
    '        Else
    '            IsLatestVersion = us.IsLatestDLLVersion(Counter.Version.ToString)
    '        End If
    '        If Not IsLatestVersion Then
    '            'Show dialog
    '            If Not IgnoreCurrentVersion Then
    '                If (New UpdateDialog).ShowDialog(us.LatestDLLVersion, us.DLLChangelog) = Windows.Forms.DialogResult.No Then
    '                    'cancel the update
    '                    Return False
    '                    Exit Function
    '                End If
    '            End If

    '            'start update

    '            'Counter = Nothing 'Release the current one, so writing to the DLL won't hurt.
    '            'Dim client As New Net.WebClient
    '            IO.File.WriteAllBytes(IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData, "mlawcdllUpdate.exe"), My.Resources.MLA_Word_Counter_DLL_Updator)
    '            IO.File.WriteAllText(IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData, "mlawcdllUpdate.exe.config"), My.Resources.MLA_Word_Counter_DLL_Updator_exe)
    '            Dim p As New Process
    '            p.StartInfo.FileName = (IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData, "mlawcdllUpdate.exe"))
    '            p.StartInfo.Arguments = """" & Environment.CurrentDirectory & """ """ & IO.Path.Combine(Environment.CurrentDirectory, "MLA Word Counter.exe") & """"
    '            p.Start()
    '            Me.Close()

    '            'ReloadCounter()
    '        Else
    '            If ShowDialogIfNoUpdate Then
    '                MessageBox.Show("There are no new updates for the counting engine.")
    '            End If
    '        End If
    '    Catch ex As TimeoutException
    '        output = False
    '        If ShowDialogIfNoUpdate Then
    '            MessageBox.Show("The check for updates timed out.  You should be able to try again in a few moments.")
    '        End If
    '    Catch ex As Exception
    '        output = False
    '        If ShowDialogIfNoUpdate Then
    '            MessageBox.Show("The check for updates failed.")
    '        End If
    '    End Try
    '    Return output
    'End Function
    Sub CloseAndSave()
        My.Settings.Content = TextBox1.Text
        Me.Close()
    End Sub
    Private Sub CopyWordCountToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopyWordCountToolStripMenuItem.Click
        My.Computer.Clipboard.SetText(Counter.GetWordCount(TextBox1.Text).ToString)
    End Sub
    Private Sub Button1_Click_2(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If String.IsNullOrEmpty(TextBox1.SelectedText) Then
            Dim s As New StatisticsWindow
            s.Show(TextBox1.Text)
        Else
            Dim s As New StatisticsWindow
            s.Show(TextBox1.SelectedText)
        End If
    End Sub
    Sub ColorHighlight()
        Dim temp As New RichTextBox
        temp.Rtf = TextBox1.Rtf

        ReCountOnSelectionChange = False

        Dim selectedIndex As Integer = TextBox1.SelectionStart
        Dim selectedLength As Integer = TextBox1.SelectionLength

        temp.SelectAll()
        temp.SelectionColor = Color.Black

        temp.Text = temp.Text.Replace(ChrW(8220), """").Replace(ChrW(8221), """") '.Replace(vbCrLf, " ")

        Dim endIndexArticle As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(Counter.ArticleAdjectiveRegexString, RegexOptions.Compiled Or RegexOptions.IgnoreCase)).Matches(temp.Text)
            If Not endIndexArticle.ContainsKey(item.Value) Then
                temp.Select(temp.Text.IndexOf(item.Value), item.Value.Length)
                endIndexArticle.Add(item.Value, temp.Text.IndexOf(item.Value) + item.Value.Length)
            Else
                temp.Select(temp.Text.IndexOf(item.Value, endIndexArticle(item.Value)), item.Value.Length)
                endIndexArticle(item.Value) = temp.Text.IndexOf(item.Value, endIndexArticle(item.Value)) + item.Value.Length
            End If
            temp.SelectionColor = Color.Gray
        Next

        Dim endIndexName As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(Counter.NameRegexString, RegexOptions.Compiled)).Matches(temp.Text)
            If Not endIndexName.ContainsKey(item.Value) Then
                temp.Select(temp.Text.IndexOf(item.Value), item.Value.Length)
                endIndexName.Add(item.Value, temp.Text.IndexOf(item.Value) + item.Value.Length)
            Else
                temp.Select(temp.Text.IndexOf(item.Value, endIndexName(item.Value)), item.Value.Length)
                endIndexName(item.Value) = temp.Text.IndexOf(item.Value, endIndexName(item.Value)) + item.Value.Length
            End If
            temp.SelectionColor = Color.Blue
        Next

        Dim endIndexNameError As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(Counter.NameErrorRegexString, RegexOptions.Compiled)).Matches(temp.Text)
            If Not endIndexNameError.ContainsKey(item.Value) Then
                temp.Select(temp.Text.IndexOf(item.Value), item.Value.Length)
                endIndexNameError.Add(item.Value, temp.Text.IndexOf(item.Value) + item.Value.Length)
            Else
                temp.Select(temp.Text.IndexOf(item.Value, endIndexNameError(item.Value)), item.Value.Length)
                endIndexNameError(item.Value) = temp.Text.IndexOf(item.Value, endIndexNameError(item.Value)) + item.Value.Length
            End If
            temp.SelectionColor = Color.Purple
        Next

        Dim endIndexDate As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(Counter.DateRegexString, RegexOptions.Compiled Or RegexOptions.IgnoreCase)).Matches(temp.Text)
            If Not endIndexDate.ContainsKey(item.Value) Then
                temp.Select(temp.Text.IndexOf(item.Value), item.Value.Length)
                endIndexDate.Add(item.Value, temp.Text.IndexOf(item.Value) + item.Value.Length)
            Else
                temp.Select(temp.Text.IndexOf(item.Value, endIndexDate(item.Value)), item.Value.Length)
                endIndexDate(item.Value) = temp.Text.IndexOf(item.Value, endIndexDate(item.Value)) + item.Value.Length
            End If
            temp.SelectionColor = Color.DarkOrange
        Next


        Dim endIndexCitation As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(Counter.CitationRegexString, RegexOptions.Compiled Or RegexOptions.IgnoreCase)).Matches(temp.Text)
            If Not endIndexCitation.ContainsKey(item.Value) Then
                temp.Select(temp.Text.IndexOf(item.Value), item.Value.Length)
                endIndexCitation.Add(item.Value, temp.Text.IndexOf(item.Value) + item.Value.Length)
            Else
                temp.Select(temp.Text.IndexOf(item.Value, endIndexCitation(item.Value)), item.Value.Length)
                endIndexCitation(item.Value) = temp.Text.IndexOf(item.Value, endIndexCitation(item.Value)) + item.Value.Length
            End If
            temp.SelectionColor = Color.Gray
        Next

        Dim endIndexQuotation As New Dictionary(Of String, Integer)
        For Each item As System.Text.RegularExpressions.Match In (New Regex(Counter.QuoteRegexString, RegexOptions.Compiled Or RegexOptions.IgnoreCase)).Matches(temp.Text)
            If Not endIndexQuotation.ContainsKey(item.Value) Then
                temp.Select(temp.Text.IndexOf(item.Value), item.Value.Length)
                endIndexQuotation.Add(item.Value, temp.Text.IndexOf(item.Value) + item.Value.Length)
            Else
                temp.Select(temp.Text.IndexOf(item.Value, endIndexQuotation(item.Value)), item.Value.Length)
                endIndexQuotation(item.Value) = temp.Text.IndexOf(item.Value, endIndexQuotation(item.Value)) + item.Value.Length
            End If
            temp.SelectionColor = Color.Red
        Next
        TextBox1.Rtf = temp.Rtf
        TextBox1.HideSelection = False
        TextBox1.Focus()
        TextBox1.Select(selectedIndex, selectedLength)
        ReCountOnSelectionChange = True
    End Sub
#Region "Loading Code"
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.Text = String.Format(Me.Text, My.Application.Info.Version.ToString)
        'TextBox1.Font = My.Settings.Font
        'FontDialog1.Font = My.Settings.Font
        TextBox1.WordWrap = My.Settings.WordWrap
        TextBox1.Text = My.Settings.Content
        CheckBox1.Checked = My.Settings.AutoCount
        My.Settings.Content = ""
        WordWrapToolStripMenuItem.Checked = My.Settings.WordWrap
        ReloadCounter()
        'DoUpdate()

        TextBox1.HideSelection = False
    End Sub
    Sub ReloadCounter()
        'If IO.File.Exists(IO.Path.Combine(Environment.CurrentDirectory, "WordCount.dll")) Then
        '    Dim a As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(IO.Path.Combine(Environment.CurrentDirectory, "WordCount.dll")) 'Load the DLL
        '    Dim types As Type() = a.GetTypes 'Get all classes in the DLL
        '    Dim plugins As New Generic.List(Of wci1_1.WordCounter) 'This will be the list of instances of your interface
        '    For Each item In types 'Look at each class in the DLL
        '        For Each intface As Type In item.GetInterfaces 'Look at each interface this class implements
        '            If intface Is GetType(wci1_1.WordCounter) Then 'Check to see if this interface is the one you are looking for
        '                plugins.Add(a.CreateInstance(item.ToString)) 'Add an instance of class implementing the interface you want to the list
        '            End If
        '        Next
        '    Next
        '    If plugins.Count = 1 Then
        '        Counter = plugins(0)
        '    ElseIf plugins.Count = 0 Then
        '        'MessageBox.Show("The files for this program are corrupt.  Please reinstall or redownload this program")
        '        'End
        '        If Not DoUpdate(True) Then
        '            MessageBox.Show("The counting engine is corrupt, and the attempt to redownload it has failed.  The program will now shut down.", "MLA Word Counter")
        '            Me.Close()
        '        End If
        '    Else
        '        'Load the first one, and ignore everything else
        '        Counter = plugins(0)
        '    End If
        'Else
        '    If Not DoUpdate(True) Then
        '        MessageBox.Show("The counting engine is missing, and the attempt to redownload it has failed.  The program will now shut down.", "MLA Word Counter")
        '        Me.Close()
        '    End If
        'End If
        Counter = New WordCount
    End Sub
#End Region
#Region "Timer Related"
    Private Sub TextBox1_SelectionChanged(sender As Object, e As System.EventArgs) Handles TextBox1.SelectionChanged
        If ReCountOnSelectionChange Then
            If Not Timer1.Enabled Then
                ColorHighlightOnTimerTick = False
                Timer1.Stop()
                Timer1.Interval = 1000
                Timer1.Start()
            End If
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        If CheckBox1.Checked AndAlso Not oldText = TextBox1.Text Then 'Otherwise the following would run if only word color has changed.
            ColorHighlightOnTimerTick = True
            Label1.Text = "Accepting input..."
            oldText = TextBox1.Text
            Timer1.Stop()
            Timer1.Interval = 1000
            Timer1.Start()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        If CheckBox1.Checked Then Count()
    End Sub
    Sub Count()
        If ColorHighlightOnTimerTick Then
            ColorHighlight()
        ElseIf CheckBox1.Checked = False Then
            ColorHighlight()
        End If

        'If GetSetting("mlawc", "keys", "indeft", False) Then
        If String.IsNullOrEmpty(TextBox1.SelectedText) Then
            Label1.Text = "Your text has " & Counter.GetWordCount(TextBox1.Text).ToString & ", not counting article adjectives, or citations." & vbCrLf & "Quotes, most dates, and most names count as one word."
        Else
            Label1.Text = "Your selection has " & Counter.GetWordCount(TextBox1.SelectedText).ToString & ", not counting article adjectives, or citations." & vbCrLf & "Quotes, most dates, and most names count as one word."
        End If
        'Else
        'Label1.Text = "Your trial of the MLA Word Counter has expired." & vbCrLf & "This version does not support activation, so please visit http://www.uniquegeeks.net/ for more information." & vbCrLf & "If you have just installed this program, please ensure that you are connected to the internet."
        'End If
    End Sub
#End Region
#Region "Menu Commands"
#Region "Edit"
    Private Sub Paste()
        TextBox1.SelectedText = My.Computer.Clipboard.GetText
    End Sub
    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        Paste()
    End Sub
#End Region
    Private Sub SelectAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        TextBox1.SelectAll()
    End Sub
    Private Sub AboutMLAWordCounterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutMLAWordCounterToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub
    Private Sub OpenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            OpenFile(OpenFileDialog1.FileName)
        End If
    End Sub
    Sub OpenFile(Filename As String)
        Dim text As String = ""
        If IO.Path.GetExtension(Filename).ToLower = ".docx" Then
            Dim doc As DocumentFormat.OpenXml.Packaging.WordprocessingDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(Filename, False)
            Dim paragraphs As Generic.List(Of DocumentFormat.OpenXml.Wordprocessing.Paragraph) = doc.MainDocumentPart.Document.Body.OfType(Of DocumentFormat.OpenXml.Wordprocessing.Paragraph).ToList
            For Each item As DocumentFormat.OpenXml.Wordprocessing.Paragraph In paragraphs
                'For Each subItem In item.ChildElements
                '    If Not subItem.GetType = GetType(DocumentFormat.OpenXml.Wordprocessing.SdtContentCitation) Then
                '        text += subItem.InnerText & vbCrLf
                '    End If
                'Next
                If item.InnerText.ToLower.Trim = "works cited" Then Exit For
                Dim out As String = item.InnerText
                Dim r1 As New System.Text.RegularExpressions.Regex("CITATION(.+?)" & Counter.CitationRegexString, RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                Dim r2 As New System.Text.RegularExpressions.Regex(Counter.CitationRegexString, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                For Each match As System.Text.RegularExpressions.Match In r1.Matches(out)
                    out = out.Replace(match.Value, " " & r2.Match(match.Value).Value)
                Next
                text += out & vbCrLf & vbCrLf
            Next
            Dim headingRegex As New System.Text.RegularExpressions.Regex(Counter.NameRegexString & "(\r\n)+" & Counter.NameRegexString & "(\r\n)+" & Counter.ClassPeriodRegexString & "(\r\n)+" & Counter.DateRegexString & "(\r\n)+" & "(.+)(\r\n)+", RegexOptions.Compiled Or RegexOptions.Multiline)
            text = headingRegex.Replace(text, "")
            Dim footer As New System.Text.RegularExpressions.Regex("(Word\sCount\:(\s)?([0-9])+(\s|\r|\n)+?)", RegexOptions.Compiled Or RegexOptions.IgnoreCase Or RegexOptions.Multiline) '?Works\sCited(\s|\r|\n)+?BIBLIOGRAPHY(.|\s|\n)+", RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Multiline)
            text = footer.Replace(text, "")

            text = text.Trim
            'text = IO.Path.GetExtension(Filename)
        Else
            'TextBox1.Text = IO.File.ReadAllText(OpenFileDialog1.FileName)
            Dim encryptedText As EncryptedText = encryptedText.FromFile(Filename)
            If encryptedText.IsValidFile Then
                If encryptedText.Encrypted Then
ShowEncryptionOptions: Dim options As New EncryptionOptions
                    If options.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Try
                            text = SymmetricEncryption.Encryptor.Decrypt(encryptedText.Content, options.Password, options.Algorythm, options.Salt, options.PaddingCharacter)
                        Catch ex As Exception
                            If MessageBox.Show("Your encryption options are incorrect.  Do you want to try again?", "Decryption Failed", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                                GoTo ShowEncryptionOptions
                            Else
                                Exit Sub
                            End If
                        End Try
                    End If
                Else
                    text = encryptedText.Content
                End If
            Else
                text = IO.File.ReadAllText(Filename)
            End If
        End If
        TextBox1.Text = text
    End Sub
    Private Sub EditFontToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        If FontDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            FontDialog1_Apply(sender, e)
        End If
    End Sub
    Private Sub FontDialog1_Apply(sender As System.Object, e As System.EventArgs) Handles FontDialog1.Apply
        TextBox1.Font = FontDialog1.Font
        My.Settings.Font = FontDialog1.Font
    End Sub
    Private Sub WordWrapToolStripMenuItem_CheckedChanged(sender As Object, e As System.EventArgs) Handles WordWrapToolStripMenuItem.CheckedChanged
        TextBox1.WordWrap = WordWrapToolStripMenuItem.Checked
        My.Settings.WordWrap = WordWrapToolStripMenuItem.Checked
    End Sub
#End Region
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Count()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Button2.Enabled = Not CheckBox1.Checked
        My.Settings.AutoCount = CheckBox1.Checked
    End Sub
End Class
