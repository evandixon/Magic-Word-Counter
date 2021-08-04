' The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

''' <summary>
''' An empty page that can be used on its own or navigated to within a Frame.
''' </summary>
Public NotInheritable Class About
    Inherits Page
    Dim _returnPage As Page
    ''' <summary>
    ''' Invoked when this page is about to be displayed in a Frame.
    ''' </summary>
    ''' <param name="e">Event data that describes how this page was reached.  The Parameter
    ''' property is typically used to configure the page.</param>
    Protected Overrides Sub OnNavigatedTo(e As Navigation.NavigationEventArgs)

    End Sub
    Public Sub New(ReturnPage As Page)
        Me.New()
        _returnPage = ReturnPage
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub GoBack(sender As Object, e As RoutedEventArgs)
        'Frame.GoBack()
        Window.Current.Content = _returnPage
    End Sub
End Class
