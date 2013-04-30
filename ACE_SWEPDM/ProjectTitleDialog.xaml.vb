Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows

Namespace AddinDialogs

    Public Class ProjectTitleDialog
        Public ShortTitle As String = ""

        ''' <summary>
        ''' On Click Of the OK Button, Validate data and Return
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
            If tbShortTitle.Text.Length > 20 Then
                System.Windows.MessageBox.Show("Short title too long, the maximum length is 20 characters", "Title Too Long...", MessageBoxButton.OK, MessageBoxImage.Warning)
            Else
                ShortTitle = tbShortTitle.Text
                DialogResult = True
                MyBase.Close()
            End If
            
        End Sub

        Private Sub Window_Activated_1(sender As Object, e As EventArgs)
            tbShortTitle.Text = ShortTitle
        End Sub
    End Class
End Namespace