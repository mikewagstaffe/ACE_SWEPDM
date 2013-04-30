Public Class ConnectionManager
    'Define the correct New function for then pre-processor directives
#If WORKMGR And ENT_PDM Then
    Public Sub New(ByVal WMConnInfo As WorkManagerConnectionInfo, ByVal EPDMConnInfo As EPDMConnectionInfo)
#ElseIf WORKMGR Then
    Public Sub New(ByVal WMConnInfo As WorkManagerConnectionInfo)
#ElseIf ENT_PDM Then
     Public Sub New( ByVal EPDMConnInfo As EPDMConnectionInfo)
#Else
    Public Sub New()
#End If

#If WORKMGR Then
        CreateWMConnection(WMConnInfo)
#End If
#If ENT_PDM Then
        CreateEPDMConnection(EPDMConnInfo)
#End If
#If DOCMGR Then
    CreateDocMgrConnection()
#End If
#If ONTOLOGYSTUDIO Then
        CreateOntologyStudioConnection()
#End If
    End Sub

#If ONTOLOGYSTUDIO Then
    Private ConnectionToOntlogyStudio As OntologyStudioConnection = Nothing

    Private Sub CreateOntologyStudioConnection()
        ConnectionToOntlogyStudio = New OntologyStudioConnection()
#If LISTEN_ONTOLOGY Then
        ConnectionToOntlogyStudio.StartWCFServer()
#End If

    End Sub
#If LISTEN_ONTOLOGY Then
     '''' <summary>
    '''' Disconnect The WCF Server EndPoint
    '''' </summary>
    '''' <remarks></remarks>
    Public Sub DisconnectWCF()
        ConnectionToOntlogyStudio.CloseWCFServer()
    End Sub
#End If

#End If

#If ENT_PDM Then
    Private ConnectionToEPDM As EPDMConnection = Nothing
    Public EPDM_ConnectionInfo As EPDMConnectionInfo = Nothing

    Private Sub CreateEPDMConnection(ByVal EPDMConnInfo As EPDMConnectionInfo)
        'Create The EPDM Connection
        EPDM_ConnectionInfo = EPDMConnInfo
        ConnectionToEPDM = New EPDMConnection(EPDM_ConnectionInfo.VaultName)
    End Sub

#End If

#If DOCMGR Then
    Private SwDocMgr As SolidWorksDocumentManager = Nothing ' used for processing documents

    Private Sub CreateDocMgrConnection()
        SwDocMgr = New SolidWorksDocumentManager()
    End Sub
#End If

#If WORKMGR Then
    Private ConnectionToWM As WorkManagerConnection = Nothing
    Public WMConnectionInfo As WorkManagerConnectionInfo = Nothing

    Private Sub CreateWMConnection(ByVal WMConnInfo As WorkManagerConnectionInfo)
        WMConnectionInfo = WMConnInfo

        If WMConnInfo IsNot Nothing Then
            'Create The Work Manager Connection
            ConnectionToWM = New WorkManagerConnection(WMConnectionInfo.ServerName, _
                                                        WMConnectionInfo.IPAddress, _
                                                        WMConnectionInfo.NetworkPort, _
                                                        WMConnectionInfo.Protocol)
            'Connect To WorkManager
            ConnectionToWM.Connect()
        End If
    End Sub

      ''' <summary>
    ''' Disconnect This Part From The Workmananager Server
    ''' </summary>
    ''' <remarks>No More Queries On the server will be possible</remarks>
    Public Sub DisconnectFromWorkManager()
        ConnectionToWM.Disconnect()
    End Sub
#End If

    ''' <summary>
    ''' Returns One of the connection objects from this manager
    ''' </summary>
    ''' <param name="Connection">Connection Object to return</param>
    ''' <returns>A connection Object</returns>
    ''' <remarks></remarks>
    Public Function GetConnection(ByVal Connection As ConnectionMgrConnection) As Object
        Select Case Connection
            Case ConnectionMgrConnection.WorkManager
#If WORKMGR Then
                Return ConnectionToWM
#Else
                Throw New NotImplementedException("Compiler Options do Not Allow Connections to Workmanager")
#End If
            Case ConnectionMgrConnection.OntologyStudio
#If ONTOLOGYSTUDIO Then
                Return ConnectionToOntlogyStudio
#Else
                Throw New NotImplementedException("Compiler Options do Not Allow Connections to Ontology Studio")
#End If
            Case ConnectionMgrConnection.EPDM
#If ENT_PDM Then
                Return ConnectionToEPDM
#Else
                Throw New NotImplementedException("Compiler Options do Not Allow Connections to Enterprise PDM")
#End If
            Case ConnectionMgrConnection.SWDocMgr
#If DOCMGR Then
                Return SwDocMgr
#Else
                Throw New NotImplementedException("Compiler Options do Not Allow Connections to the Document Manager")
#End If
            Case Else
                Return Nothing
        End Select
    End Function
End Class

Public Enum ConnectionMgrConnection
    WorkManager
    OntologyStudio
    EPDM
    SWDocMgr
End Enum
