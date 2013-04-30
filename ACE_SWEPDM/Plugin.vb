' (C) Copyright 2013 by M.Wagstaffe (Ishida Europe Ltd)

'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Windows

Imports ItemGatewayDataBaseComs.API

' This line is not mandatory, but improves loading performances
<Assembly: ExtensionApplication(GetType(ACE_SWEPDM.Plugin))> 

Namespace ACE_SWEPDM

    ' This class is instantiated by AutoCAD once and kept alive for the 
    ' duration of the session. If you don't do any one time initialization 
    ' then you should remove this class.
    Public Class Plugin
        Implements IExtensionApplication

        Private DocumentManager As DocumentCollection = Nothing

        Public Sub Initialize() Implements IExtensionApplication.Initialize
            Try
                'Dont Want this yet
                'DocumentManager = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
                'AddHandler DocumentManager.DocumentCreated, AddressOf callback_documentCreated


                'Load The Settings From The Central Settings Dataabse
                Dim sEPDMName As String = ItemGatewayDatabase.GetSetting("EPDMName")
                Dim sManagedDocumentsFolder As String = ItemGatewayDatabase.GetSetting("ManagedDocumentsFolder")
                Dim sBomLayoutName As String = ItemGatewayDatabase.GetSetting("BomLayoutName")
                'Create The Connection Manager
                ACECommands.ConnectionMgr = New ConnectionManager(New EPDMConnectionInfo(sEPDMName, sManagedDocumentsFolder, sBomLayoutName))
                If Autodesk.Windows.ComponentManager.Ribbon Is Nothing Then
                    'The ribbon is not available so register an event to wait for it
                    AddHandler Autodesk.Windows.ComponentManager.ItemInitialized, AddressOf ComponentManager_ItemInitialised
                Else
                    'was loaded by netload so ribbon is available
                    LoadRibbon()
                End If
            Catch ex As System.Exception
                System.Windows.MessageBox.Show("Critical Error, Initialising ACE_SWEPDM Addin." & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, "ACE_SWEPDM Fatal Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            End Try
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
            RemoveHandler DocumentManager.DocumentCreated, AddressOf callback_documentCreated
            ConnectionMgr = Nothing
        End Sub

        ''' <summary>
        ''' Handler For The Document Created Event
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub callback_documentCreated(ByVal sender As Object, ByVal e As DocumentCollectionEventArgs)
            If e.Document = Nothing Then
                Exit Sub
            End If

            Dim db As Database = e.Document.Database
            AddHandler db.BeginSave, AddressOf callback_BeginSave
            AddHandler db.SaveComplete, AddressOf callback_SaveComplete
        End Sub


        ''' <summary>
        ''' Handler for the start of a document Save Event
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub callback_BeginSave(ByVal sender As Object, ByVal e As DatabaseIOEventArgs)
            Dim myDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        End Sub

        ''' <summary>
        ''' Handler for the document saved event
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub callback_SaveComplete(ByVal sender As Object, ByVal e As DatabaseIOEventArgs)
            Dim myDwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        End Sub

        ''' <summary>
        ''' Called when The Ribon is initialised, this is used to autolad the ribbon interface
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Sub ComponentManager_ItemInitialised(ByVal sender As Object, ByVal e As RibbonItemEventArgs)
            If Autodesk.Windows.ComponentManager.Ribbon IsNot Nothing Then
                LoadRibbon()
                'Ribbon is available remove this event handler to prevent events
                RemoveHandler Autodesk.Windows.ComponentManager.ItemInitialized, AddressOf ComponentManager_ItemInitialised
            End If
        End Sub

        ''' <summary>
        ''' Load The Ribbon Interface
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub LoadRibbon()
            Dim EPDMRibbonControl As RibbonControl = ComponentManager.Ribbon 'New RibbonControl()
            Dim EPDMRibbonTab As RibbonTab = New RibbonTab()
            EPDMRibbonTab.Title = "Enterprise PDM"
            EPDMRibbonTab.Id = "PDM_TAB_ID"
            EPDMRibbonControl.Tabs.Add(EPDMRibbonTab)

            Dim ProjectToolsPanel As RibbonPanelSource = New RibbonPanelSource()
            ProjectToolsPanel.Title = "Project Tools"
            Dim ProjectPanel As RibbonPanel = New RibbonPanel()
            ProjectPanel.Source = ProjectToolsPanel

            Dim PartToolsPanel As RibbonPanelSource = New RibbonPanelSource()
            PartToolsPanel.Title = "Part Tools"
            Dim PartPanel As RibbonPanel = New RibbonPanel()
            PartPanel.Source = PartToolsPanel

            Dim PDMToolsPanel As RibbonPanelSource = New RibbonPanelSource()
            PDMToolsPanel.Title = "PDM Tools"
            Dim PDMPanel As RibbonPanel = New RibbonPanel()
            PDMPanel.Source = PDMToolsPanel

            Dim MiscToolsPanel As RibbonPanelSource = New RibbonPanelSource()
            MiscToolsPanel.Title = "Miscellaneous"
            Dim MiscPanel As RibbonPanel = New RibbonPanel()
            MiscPanel.Source = MiscToolsPanel

            EPDMRibbonTab.Panels.Add(ProjectPanel)
            EPDMRibbonTab.Panels.Add(PartPanel)
            EPDMRibbonTab.Panels.Add(PDMPanel)
            EPDMRibbonTab.Panels.Add(MiscPanel)

            Dim CreateProjectButton As RibbonButton = New RibbonButton()
            CreateProjectButton.Text = "Create New" & vbCrLf & "Project"
            CreateProjectButton.ShowText = True
            CreateProjectButton.ShowImage = True
            CreateProjectButton.Image = LoadImage(My.Resources.projnew)
            CreateProjectButton.LargeImage = LoadImage(My.Resources.projnew)
            CreateProjectButton.Orientation = System.Windows.Controls.Orientation.Vertical
            CreateProjectButton.Size = RibbonItemSize.Large
            CreateProjectButton.CommandParameter = "EPDM_CREATEPROJECT"
            CreateProjectButton.CommandHandler = New RibbonCommandHandler()

            Dim CreateDrawingButton As RibbonButton = New RibbonButton()
            CreateDrawingButton.Text = "Create New" & vbCrLf & "Drawings"
            CreateDrawingButton.ShowText = True
            CreateDrawingButton.ShowImage = True
            CreateDrawingButton.Image = LoadImage(My.Resources.documentnew)
            CreateDrawingButton.LargeImage = LoadImage(My.Resources.documentnew)
            CreateDrawingButton.Orientation = System.Windows.Controls.Orientation.Vertical
            CreateDrawingButton.Size = RibbonItemSize.Large
            CreateDrawingButton.CommandParameter = "EPDM_NEWDRAWINGS"
            CreateDrawingButton.CommandHandler = New RibbonCommandHandler()

            Dim LockProjectButton As RibbonButton = New RibbonButton()
            LockProjectButton.Text = "Check Out" & vbCrLf & "Project"
            LockProjectButton.ShowText = True
            LockProjectButton.ShowImage = True
            LockProjectButton.Image = LoadImage(My.Resources.locked)
            LockProjectButton.LargeImage = LoadImage(My.Resources.locked)
            LockProjectButton.Orientation = System.Windows.Controls.Orientation.Vertical
            LockProjectButton.Size = RibbonItemSize.Large
            LockProjectButton.CommandParameter = "EPDM_LOCK"
            LockProjectButton.CommandHandler = New RibbonCommandHandler()

            Dim UnlockProjectButton As RibbonButton = New RibbonButton()
            UnlockProjectButton.Text = "Check In" & vbCrLf & "Project"
            UnlockProjectButton.ShowText = True
            UnlockProjectButton.ShowImage = True
            UnlockProjectButton.Image = LoadImage(My.Resources.unlocked)
            UnlockProjectButton.LargeImage = LoadImage(My.Resources.unlocked)
            UnlockProjectButton.Orientation = System.Windows.Controls.Orientation.Vertical
            UnlockProjectButton.Size = RibbonItemSize.Large
            UnlockProjectButton.CommandParameter = "EPDM_UNLOCK"
            UnlockProjectButton.CommandHandler = New RibbonCommandHandler()

            Dim ChangeStateProjectButton As RibbonButton = New RibbonButton()
            ChangeStateProjectButton.Text = "Change" & vbCrLf & "State"
            ChangeStateProjectButton.ShowText = True
            ChangeStateProjectButton.ShowImage = True
            ChangeStateProjectButton.Image = LoadImage(My.Resources.commit)
            ChangeStateProjectButton.LargeImage = LoadImage(My.Resources.commit)
            ChangeStateProjectButton.Orientation = System.Windows.Controls.Orientation.Vertical
            ChangeStateProjectButton.Size = RibbonItemSize.Large
            ChangeStateProjectButton.CommandParameter = "EPDM_CHANGESTATE"
            ChangeStateProjectButton.CommandHandler = New RibbonCommandHandler()

            ProjectToolsPanel.Items.Add(CreateProjectButton)
            ProjectToolsPanel.Items.Add(New RibbonSeparator())
            ProjectToolsPanel.Items.Add(CreateDrawingButton)
            PDMToolsPanel.Items.Add(LockProjectButton)
            PDMToolsPanel.Items.Add(UnlockProjectButton)
            PDMToolsPanel.Items.Add(ChangeStateProjectButton)

            'EPDMRibbonTab.IsActive = True
        End Sub

        ''' <summary>
        ''' Loads a PNG image from the Plugins Resource File, and Returns a BitMapImage Object For this image
        ''' </summary>
        ''' <param name="pic">image name of an image in the resource library, to load </param>
        ''' <returns>BitmapImage of the image selected</returns>
        ''' <remarks></remarks>
        Private Function LoadImage(ByVal pic As System.Drawing.Bitmap) As Windows.Media.Imaging.BitmapImage
            Dim ms As IO.MemoryStream = New IO.MemoryStream()
            pic.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
            Dim bi As Windows.Media.Imaging.BitmapImage = New Windows.Media.Imaging.BitmapImage()
            bi.BeginInit()
            bi.StreamSource = ms
            bi.EndInit()
            Return bi
        End Function


    End Class


    Public Class RibbonCommandHandler
        Implements System.Windows.Input.ICommand

        Public Function CanExecute(ByVal parameter As Object) As Boolean Implements System.Windows.Input.ICommand.CanExecute
            Return True
        End Function

        Public Event CanExecuteChanged(ByVal sender As Object, ByVal e As System.EventArgs) Implements System.Windows.Input.ICommand.CanExecuteChanged

        Public Sub Execute(ByVal parameter As Object) Implements System.Windows.Input.ICommand.Execute

            If TypeOf parameter Is RibbonButton Then
                Dim button As RibbonButton = TryCast(parameter, RibbonButton)
                If button IsNot Nothing Then
                    Select Case DirectCast(button.CommandParameter, String)
                        Case "EPDM_CREATEPROJECT"
                            ACECommands.NewElectricalProject()
                            Exit Select
                    End Select
                    'Old code to send command to autocad now just run the acual function
                    ' Dim app As Autodesk.AutoCAD.Interop.AcadApplication = DirectCast(Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
                    'app.ActiveDocument.SendCommand(DirectCast(button.CommandParameter, String))
                End If
            End If
        End Sub
    End Class
End Namespace