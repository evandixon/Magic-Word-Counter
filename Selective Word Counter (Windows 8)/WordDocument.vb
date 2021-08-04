'Public Class WordDocument
'    Public Shared Function OpenWordFile()
'        Dim localFolder = Windows.Storage.ApplicationData.Current.LocalFolder

'        Dim file = Await localFolder.GetFileAsync(Test.zip)

'        Dim randomStream = Await file.OpenReadAsync()



'        Using stream As Stream = randomStream.AsStreamForRead()


'            Dim zipArchive = New System.IO.Compression.ZipArchive(stream)



'            ' Entries contains all the files in the ZIP

'            For Each entry As var In zipArchive.Entries


'                Using entryStream = entry.Open()


'                    '

'                    ' Read string file content

'                    '

'                    'using (var streamReader = new StreamReader(entryStream))

'                    '{

'                    '    var result = await streamReader.ReadToEndAsync();

'                    '}



'                    '

'                    ' Read string file content (using WinRT APIs)

'                    '

'                    Using inputStream = entryStream.AsInputStream()


'                        Using reader As New DataReader(inputStream)


'                            Dim fileSize = Await reader.LoadAsync(CUInt(entryStream.Length))




'                            ' You can also use the other methods available on DataReader to load an IBuffer, Byte Array, etc.

'                            Dim stringContent = reader.ReadString(fileSize)

'                        End Using

'                    End Using

'                End Using

'            Next
'        End Using

'        Using package__1 = Package.Open("generated.docx")
'            Dim part = package__1.GetPart(New Uri("/word/document.xml", UriKind.Relative))

'            Dim stream = part.GetStream()

'            Dim xmlDocument = New XmlDocument()

'            xmlDocument.Load(stream)

'            ' do your changes here

'            stream.Seek(0, SeekOrigin.Begin)

'            xmlDocument.Save(stream)

'            stream.SetLength(stream.Position)
'        End Using

'    End Function
'End Class
