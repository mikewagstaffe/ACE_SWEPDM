Imports System.Runtime.InteropServices

Module WindowsAPI

    ''' <summary>
    ''' Uses the Windows API to bring the SldWorks process into focus
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BringSolidWorksToFront()
        BringApplicationToFront("SldWorks")
    End Sub

    Public Sub MinimiseOntologyStudio()
        MinimiseApplication("OntologyStudioClient")
    End Sub

    Private Sub BringApplicationToFront(ByVal ApplicationName As String)
        Try
            Dim ProcessName As String = ApplicationName
            Dim ExistingInstances() As Process = Process.GetProcessesByName(ProcessName)
            If ExistingInstances.Length > 0 Then
                ShowWindow(ExistingInstances.First.MainWindowHandle, ShowWindowCommands.SW_FORCEMINIMIZE)
                'ShowWindow(ExistingInstances.First.MainWindowHandle, ShowWindowCommands.SW_RESTORE)
                ShowWindow(ExistingInstances.First.MainWindowHandle, ShowWindowCommands.SW_SHOWNORMAL)
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub MinimiseApplication(ByVal ApplicationName As String)
        Try
            Dim ProcessName As String = ApplicationName
            Dim ExistingInstances() As Process = Process.GetProcessesByName(ProcessName)
            If ExistingInstances.Length > 0 Then
                ShowWindow(ExistingInstances.First.MainWindowHandle, ShowWindowCommands.SW_FORCEMINIMIZE)
            End If
        Catch ex As Exception
        End Try
    End Sub

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Function ShowWindow(ByVal hwnd As IntPtr, ByVal nCmdShow As ShowWindowCommands) As Boolean
    End Function

    Enum ShowWindowCommands As Integer
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
    End Enum

End Module
