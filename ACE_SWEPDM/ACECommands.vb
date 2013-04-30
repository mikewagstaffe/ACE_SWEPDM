Imports System.ComponentModel
Imports System.Threading
Imports ItemGatewayDataBaseComs.API

Module ACECommands
    Public ConnectionMgr As ConnectionManager = Nothing

    Public CurrentProgress As Integer = 0
    Public ProgressMessage As String = ""
    Private trd As Thread = Nothing

    Public Sub NewElectricalProject()
        Dim NewACEProject As ManagedACEProject = New ManagedACEProject(ConnectionMgr, True)
    End Sub
End Module