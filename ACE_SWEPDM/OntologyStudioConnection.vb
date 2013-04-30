Imports WCFComms
Imports System.ServiceModel

Public Class OntologyStudioConnection

    ''' <summary>
    ''' Creates A New Item in Ontology Studio
    ''' </summary>
    ''' <returns>Ontology Studio Info Detailing The Clasification of the item</returns>
    ''' <remarks></remarks>
    Public Function ClassifyNewItem(Optional ByVal ItemNumber As String = "", _
                                  Optional ByVal String1 As String = "", _
                                  Optional ByVal String2 As String = "", _
                                  Optional ByVal String3 As String = "") As OntologyStudioInfo
        Try
            Dim SendToOSInfo As OntologyStudioInfo = New OntologyStudioInfo()
            If String1 = "" And String2 = "" Then
                String2 = "New Part From Item Gateway"
            End If
            SendToOSInfo.Item = ItemNumber
            SendToOSInfo.sDataString1 = String1
            SendToOSInfo.sDataString2 = String2
            SendToOSInfo.sDataString3 = String3
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(SendToOSInfo, WCFComms.ItemMode.ModeNormal)
            Return OsInfo
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Perform A Query On Ontology Studio to obtain the Ontology Studio Allocated Partnumber
    ''' </summary>
    ''' <param name="Name">Part Number to check</param>
    ''' <returns>Returns the Ontology Studio Allocated Partnumber</returns>
    ''' <remarks></remarks>
    Public Function LookupInOntologyStudio(ByVal Name As String) As String
        Try
            Dim OsItemInfo As OntologyStudioInfo = New OntologyStudioInfo()
            OsItemInfo.Item = Name
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(OsItemInfo, ItemMode.ModeSilent)

            Return OsInfo.Item
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Requests Part Data From Ontology Studio
    ''' </summary>
    ''' <param name="Path">Path To The File</param>
    ''' <param name="Name">Name of the item</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RequestPartData(ByRef Path As String, ByRef Name As String) As Boolean
        Try
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(New OntologyStudioInfo(), ItemMode.ModeLookup)
            Path = OsInfo.sDataString2
            Name = OsInfo.Item
            Return True
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' <summary>
    ''' Perform A Query On Ontology Studio to obtain the clasification info
    ''' </summary>
    ''' <param name="Name">Part Number to check</param>
    ''' <returns>Returns the Ontology Studio Allocated Partnumber</returns>
    ''' <remarks></remarks>
    Public Function GetClasificationInfo(ByVal Name As String) As OntologyStudioInfo
        Try
            Dim OsItemInfo As OntologyStudioInfo = New OntologyStudioInfo()
            OsItemInfo.Item = Name
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(OsItemInfo, ItemMode.ModeSilent)

            Return OsInfo
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Checks To See if a Part Exists In Ontology Studio
    ''' </summary>
    ''' <param name="Name">The Ontology Studio Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExistsInOntologyStudio(ByVal Name As String) As Boolean
        Try
            Dim OsItemInfo As OntologyStudioInfo = New OntologyStudioInfo()
            OsItemInfo.Item = Name
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(OsItemInfo, ItemMode.ModeSilent)

            If OsInfo.Item Is Nothing Then Return False

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Clasify A Part in Ontology studio
    ''' </summary>
    ''' <param name="ClasificationInfo">Clasification Information</param>
    ''' <returns>0 If Succsefull</returns>
    ''' <remarks></remarks>
    Public Function Clasify(ByRef ClasificationInfo As OntologyStudioData, Optional ByVal NewPart As Boolean = False) As Integer
        Try
            Dim SendToOSInfo As OntologyStudioInfo = New OntologyStudioInfo()
            SendToOSInfo.Item = ClasificationInfo.OriginalName
            SendToOSInfo.sDataString1 = ClasificationInfo.AssociatedFileName
            SendToOSInfo.sDataString2 = ClasificationInfo.OriginalDescription
            SendToOSInfo.sDataString3 = ClasificationInfo.ScreenShotImgPath
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(SendToOSInfo, WCFComms.ItemMode.ModeNormal)
            If OsInfo.Item Is Nothing Then
                Return 1
            End If

            If OsInfo.Result = 30 Then
                ' the result returned is an alternative part to use
                Return 2
            End If
            ClasificationInfo.Name = OsInfo.Item
            ClasificationInfo.RelativePath = OsInfo.sDataString2
            ClasificationInfo.Description = OsInfo.sDataString1

            If String.Compare(ClasificationInfo.Name, ClasificationInfo.OriginalName, True) <> 0 Then
                ClasificationInfo.RenameRequired = True
            Else
                ClasificationInfo.RenameRequired = False
            End If
            Return 0
        Catch ex As Exception
            Throw New Exception("Failed To Clasify A Part", ex)
        End Try
    End Function

    ''' <summary>
    ''' Looks Up a Given Part in Ontology Studio
    ''' </summary>
    ''' <param name="UniqueID">Ontology Studio ID of the Part to lookup</param>
    ''' <remarks></remarks>
    Public Sub FindInOntologyStudio(ByVal UniqueID As String)
        Try
            Dim SendToOSInfo As OntologyStudioInfo = New OntologyStudioInfo()
            SendToOSInfo.Item = UniqueID
            SendToOSInfo.sDataString1 = ""
            SendToOSInfo.sDataString2 = ""
            SendToOSInfo.sDataString3 = ""
            Dim OsInfo As OntologyStudioInfo = SendCommandToOntologyStudio(SendToOSInfo, WCFComms.ItemMode.ModeShow)
        Catch ex As Exception
            Throw New Exception("Failed To Loook Up Part", ex)
        End Try
    End Sub

    'todo fetch from central ItemGateway database
    Private ReadOnly ONTOLOGYSTUDIONAME As String = "Ontology Studio Client"
    Private ReadOnly ONTOLOGYSTUDIOEXE As String = "Ontology Studio Client.exe"

    ''' <summary>
    ''' The pipeProxy object is the access channel used to call exported commands in ItemGateway
    ''' </summary>
    Private pipeProxy As IOntologyStudioCommand = Nothing
    ''' <summary>
    ''' The pipeFactory object is used to connect the pipeProxy to the OntologyStudio command server
    ''' </summary>
    ''' <remarks></remarks>
    Private pipeFactory As ChannelFactory(Of IOntologyStudioCommand) = Nothing

    ''' <summary>
    ''' Connects the the Ontology Studio WCF command server 
    ''' </summary>
    Private Sub ConnectToOntologyStudio()
        Try
            Dim namedPipe As NetNamedPipeBinding = New NetNamedPipeBinding()
            namedPipe.SendTimeout = New TimeSpan(1, 5, 0)
            pipeFactory = New ChannelFactory(Of IOntologyStudioCommand)(namedPipe, New EndpointAddress("net.pipe://localhost/OntologyStudio/PipeCommandFunction"))

            If pipeFactory IsNot Nothing Then
                pipeProxy = pipeFactory.CreateChannel()
            Else
                Throw New Exception("Failed to create Connection to Ontology Studio Function Server")
            End If
            If pipeProxy Is Nothing Then
                Throw New Exception("Failed to open channel to Ontology Studio Function Server")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' closes the current open connection to the Ontology Studio WCF command server
    ''' </summary>
    Private Sub DisconnectOntologyStudio()
        Try
            pipeProxy = Nothing
            pipeFactory.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Reads Executable path and name from the registry. These are set on a user basis by the Ontology Studio Client install
    ''' </summary>
    ''' <param name="OntologyStudioName"></param>
    ''' <param name="OntologyStudioExe"></param>
    ''' <returns>True if success + OntologyStudioName set and OntologyStudioExe set, False if fails</returns>
    ''' <remarks></remarks>
    Private Function ReadExePropertiesFromRegistry(ByRef OntologyStudioName As String, ByRef OntologyStudioExe As String) As Boolean
        Try
            Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Ontology Studio", False)
            Dim Name As String = regKey.GetValue("Name").ToString()
            Dim Executable As String = regKey.GetValue("Executable").ToString()
            regKey.Close()
            If Name Is Nothing Then Return False
            OntologyStudioName = Name
            If Executable Is Nothing Then Return False
            OntologyStudioExe = Executable
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Sends a correctly formatted command to the Ontology Studio WCF command server
    ''' </summary>
    ''' <param name="Item">The identification string of the item to work with</param>
    ''' <param name="Mode">an optional operating mode, some commands canot operate in all modes, 
    ''' mode will be ignored in this case</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SendCommandToOntologyStudio(ByVal Item As OntologyStudioInfo, Optional ByVal Mode As ItemMode = ItemMode.ModeNormal) As OntologyStudioInfo
        Try
            ConnectToOntologyStudio()
            Dim Retval As OntologyStudioInfo = pipeProxy.Process(Item, Mode)
            DisconnectOntologyStudio()
            Return Retval
        Catch ex As EndpointNotFoundException

            Dim OntologyStudioName As String = ""
            Dim OntologyStudioExe As String = ""
            If Not ReadExePropertiesFromRegistry(OntologyStudioName, OntologyStudioExe) Then Return Nothing

            'The endpoint was not found this means that the server is probably not running
            If IsProcessRunning(OntologyStudioName) Then
                Throw ex
            End If
            Try
                Dim OntologyStudioProcess As Process = Process.Start(OntologyStudioExe)

                Dim SleepCount As Integer = 60
                While True
                    If OntologyStudioProcess.Responding Then
                        Exit While
                    End If
                    System.Threading.Thread.Sleep(1000)
                    SleepCount -= 1
                    If SleepCount <= 0 Then
                        Throw New Exception()
                    End If
                End While
                'allowtime to make sure that server is started
                System.Threading.Thread.Sleep(5000)
                Try
                    ConnectToOntologyStudio()
                    Dim Retval As OntologyStudioInfo = pipeProxy.Process(Item, Mode)
                    DisconnectOntologyStudio()
                    Return Retval
                Catch exagain As Exception
                    Throw ex
                End Try
            Catch startex As Exception
                Throw startex
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private ThreadParam_Item As OntologyStudioInfo = Nothing
    Private ThreadParam_Mode As ItemMode = ItemMode.ModeSilent
    Private ThreadReturn As OntologyStudioInfo = Nothing

    Private Sub OntologyStudioRequest()
        ThreadReturn = SendCommandToOntologyStudio(ThreadParam_Item, ThreadParam_Mode)
    End Sub
    ''' <summary>
    ''' Checks to see if the given process name is currently running
    ''' </summary>
    ''' <param name="ProcessName">name of the process to find</param>
    ''' <returns>True if the process is running</returns>
    Private Function IsProcessRunning(ByVal ProcessName As String) As Boolean
        Dim Processes() As Process = Process.GetProcesses()
        For Each Proc As Process In Processes
            If Proc.ProcessName.ToUpper() = ProcessName.ToUpper() Then
                Return True
            End If
        Next
        Return False
    End Function
#If LISTEN_ONTOLOGY Then

    Private host As ServiceHost = New ServiceHost(GetType(ItemGatewayCommand), New Uri("net.pipe://localhost/ItemGateway"))

    ''' <summary>
    ''' Start a WCF Server Running To Listen For Incoming Requests
    ''' </summary>
    ''' <returns>True If started</returns>
    ''' <remarks></remarks>
    Public Function StartWCFServer() As Boolean
        Try
            host.AddServiceEndpoint(GetType(IItemGatewayCommand), New NetNamedPipeBinding(), "PipeCommandFunction")
            If host Is Nothing Then
                Throw New Exception("Failed to start the WCF command server")
            End If
            host.Open()
            Dim SleepCount As Integer = 6
            While True
                If host.State = CommunicationState.Opened Then
                    Return True
                End If
                SleepCount -= 1
                If SleepCount <= 0 Then
                    Return False
                End If
                System.Threading.Thread.Sleep(200)
            End While
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Closes the WCF command server for Ontology Studio
    ''' </summary>
    ''' <returns>true if succesfully closed</returns>
    Public Function CloseWCFServer() As Boolean
        Try
            If Not host.State = CommunicationState.Closed Then
                host.Close()
            End If
            Dim SleepCount As Integer = 6
            While True
                If host.State = CommunicationState.Closed Then
                    Return True
                End If
                SleepCount -= 1
                If SleepCount <= 0 Then
                    Return False
                End If
                System.Threading.Thread.Sleep(200)
            End While
            Return False

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End If
End Class
