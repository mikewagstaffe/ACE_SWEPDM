Public Class EPDMConnectionInfo
    Private _VaultName As String = Nothing
    Private _ManagedDocumentsFolder As String = Nothing
    Private _BomLayoutName As String = Nothing

    Public Sub New(ByVal VaultName As String, ByVal ManagedDocumentsFolder As String, ByVal BomLayoutName As String)
        _VaultName = VaultName
        _ManagedDocumentsFolder = ManagedDocumentsFolder
        _BomLayoutName = BomLayoutName
    End Sub

    Public ReadOnly Property VaultName() As String
        Get
            Return _VaultName
        End Get
    End Property

    Public ReadOnly Property ManagedDocumentsFolder() As String
        Get
            Return _ManagedDocumentsFolder
        End Get
    End Property

    Public ReadOnly Property BomLayoutName() As String
        Get
            Return _BomLayoutName
        End Get
    End Property

End Class
