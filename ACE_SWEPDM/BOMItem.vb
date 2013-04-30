Public Class BOMItem
    Public PartNumber As String = ""
    Public Issue As String = ""
    Public Quantity As Integer = 0
    Public ElementID As String = ""
    Public IsClassifed As Boolean = False

    Public Sub New(ByVal PartNumber As String, ByVal Issue As String, ByVal Quantity As String, ByVal ElementID As String)
        Me.PartNumber = PartNumber
        Me.Issue = Issue
        Integer.TryParse(Quantity, Me.Quantity)
        Me.ElementID = ElementID
    End Sub

    Public Sub New(ByVal PartNumber As String, ByVal Quantity As Integer)
        Me.PartNumber = PartNumber
        Me.Quantity = Quantity
    End Sub
    Public Function Copy() As BOMItem
        Dim CopyItem As BOMItem = New BOMItem(Me.PartNumber, Me.Issue, Me.Quantity.ToString(), Me.ElementID)
        CopyItem.IsClassifed = Me.IsClassifed
        Return CopyItem
    End Function
End Class
