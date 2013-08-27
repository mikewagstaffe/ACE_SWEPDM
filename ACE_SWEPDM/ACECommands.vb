Imports System.ComponentModel
Imports System.Threading
Imports ItemGatewayDataBaseComs.API
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices

Module ACECommands
    Private Const csErrorLogFile = "ACADE_EPPD_Error.Log"
    Private Const ciMaxLogLevel = 5                             'The Maximum Number of Exceptions to log when itterating down a tree
    Public ConnectionMgr As ConnectionManager = Nothing
    Public UserFolder As String = ""
    Public CurrentProgress As Integer = 0
    Public ProgressMessage As String = ""
    Private trd As Thread = Nothing

    ''' <summary>
    ''' Creates A New AutoCad Electrical Project, From Scratch
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub NewElectricalProject()
        Try
            Dim NewACEProject As ManagedACEProject = New ManagedACEProject(ConnectionMgr, True)
            If NewACEProject IsNot Nothing Then
                System.Windows.MessageBox.Show("Project was succesfully created.", "Create Project", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Information)
            Else
                System.Windows.MessageBox.Show("A fatal error occured creating the new project.", "Creat Project Error", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
            End If
        Catch ex As System.Exception
            ReportAndLogError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Copies An Existing AutoCAD Electrical Project
    ''' </summary>
    ''' <remarks>Either Copies the Current Acive Project or a project from Ontology Studio</remarks>
    Public Sub NewCopyOfElectricalProject()
        Try
            Dim ExistingProject As ManagedACEProject = New ManagedACEProject(ConnectionMgr) 'Create An Object To Represent The Existing Project

            'First Prompt if The Active Project Is To Be used as the source project
            Dim mbResult As Windows.MessageBoxResult = System.Windows.MessageBox.Show("Would you like to make a copy of the active project, or select a project from Ontology Studio?", "Source Project", Windows.MessageBoxButton.YesNo, Windows.MessageBoxImage.Question, Windows.MessageBoxResult.No)
            If mbResult = Windows.MessageBoxResult.Yes Then
                'Find out what the active project is
                If Not ExistingProject.SetACEActiveProject() Then Return
            Else
                'request a project number from ontolgy studio
                If Not ExistingProject.SetProject() Then Return
            End If
            If ExistingProject.IsProjectDrawingOpen() Then
                Dim mbResultOpenDwg As Windows.MessageBoxResult = System.Windows.MessageBox.Show("The source project has open drawings. Should these be closed automatically, all changes will be saved?", "Source Project Has Open Drawings", Windows.MessageBoxButton.YesNo, Windows.MessageBoxImage.Question, Windows.MessageBoxResult.No)
                If mbResult = Windows.MessageBoxResult.No Then
                    System.Windows.MessageBox.Show("Source project has open drawings, cannot continue.", "Source Project Has Open Drawings", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
                    Return
                End If
            End If
            Dim NewACEProject As ManagedACEProject = Nothing
            Select Case ExistingProject.MakeCopyProject(NewACEProject)
                Case 0
                    System.Windows.MessageBox.Show("Project was succesfully copied.", "Copy Project", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Information)
                    Exit Select
                Case -1
                    System.Windows.MessageBox.Show("Invalid source project selected.", "Copy Project Error", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
                    Exit Select
                Case -2
                    System.Windows.MessageBox.Show("Source Project Had Open Drawings which could not be closed.", "Copy Project Error", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
                    Exit Select
                Case Else
                    System.Windows.MessageBox.Show("A fatal error occured copying source project.", "Copy Project Error", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
                    Exit Select
            End Select
        Catch ex As System.Exception
            ReportAndLogError(ex)
        End Try
    End Sub

    

    Private Sub ReportAndLogError(ByVal Ex As System.Exception)
        Try
            If UserFolder.Length > 0 Then 'Only log to a file if there is somewhere to put it
                If Not UserFolder.EndsWith("\") Then UserFolder &= "\"
                Dim LogFile As String = UserFolder & csErrorLogFile
                Dim LogWriter As New System.IO.StreamWriter(LogFile, True)
                Dim ExCount As Integer = 0
                Dim LogEx As System.Exception = Ex
                LogWriter.WriteLine("Error Log Entry: " & DateTime.Now.ToShortDateString() & " @ " & DateTime.Now.ToShortTimeString())
                While ExCount < ciMaxLogLevel
                    ExCount += 1
                    LogWriter.WriteLine(LogEx.Message)
                    LogWriter.WriteLine(LogEx.StackTrace)
                    If LogEx.InnerException IsNot Nothing Then
                        LogEx = LogEx.InnerException
                    End If
                End While
                LogWriter.Flush()
                LogWriter.Close()
                LogWriter = Nothing
            End If
        Catch
            'No point report this error, this is a fundamental error if we cannot log error information
        End Try
        System.Windows.MessageBox.Show("A serious error occure whilst carrying out the requested command." & vbCrLf & _
                                          "Report this error to your system Administrator.", "Fatal Error", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
    End Sub
End Module