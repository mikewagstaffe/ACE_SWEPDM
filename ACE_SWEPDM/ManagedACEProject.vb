Imports System.Xml
Imports ItemGatewayDataBaseComs.API
Imports System.ComponentModel
Imports Autodesk.Windows
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Xml.XPath

Public Class ManagedACEProject

    Private Const csClasificationImage As String = "AcadProjectImg.png"
    Private ClassificationInfo As OntologyStudioData = Nothing
    Private ConnectionMgr As ConnectionManager = Nothing
    Private EPDMItemID As Integer = -1
    Private EPDMFileID As Integer = -1
    Private ProjectInfoFile As String = Nothing
    Private PartNumberPattern As String = "*"
    Private ProjectTitle As String = ""
    Private ProjectFolder As String = ""
    Private ReportProgress As Boolean = False
    Private ProgressWorker As BackgroundWorker = Nothing
    Private CurrentOperationProgress As Integer = 0
    Private ProjectExists As Boolean = False
    Private DrawingsList As List(Of String) = Nothing
    Private FullProjectName As String = ""
    ''' <summary>
    ''' Creates A new Manaegd ACE Project Object
    ''' </summary>
    ''' <param name="bCreate">if True A new project is created in the Managed system, else 
    ''' The Object can be used to represent an exiting Project </param>
    ''' <remarks></remarks>
    Public Sub New(ByRef ConnectionMgr As ConnectionManager, Optional ByVal bCreate As Boolean = False)
        Me.ConnectionMgr = ConnectionMgr
        If bCreate Then
            CreateNewProject()
        End If
    End Sub

    ''' <summary>
    ''' Creates a new Managed Project structure in all systems
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateNewProject()
        Try
            PartNumberPattern = ItemGatewayDatabase.GetSetting("PartNumberPattern")
            ClassificationInfo = New OntologyStudioData()
            ClassificationInfo.OriginalName = "</Pattern>" & PartNumberPattern
            ClassificationInfo.AssociatedFileName = "New AutoCAD Electrical Project"

            Dim sAppPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
            Dim sMasterImgPath As String = ""
            If sAppPath IsNot Nothing Then
                If sAppPath.StartsWith("FILE:\", True, Nothing) Then
                    sMasterImgPath = sAppPath.Substring(6)
                Else
                    sMasterImgPath = sAppPath
                End If

                If Not sMasterImgPath.EndsWith("\") Then
                    sMasterImgPath &= "\"
                End If
                sMasterImgPath &= csClasificationImage

                'If the master file exists then copy it
                If FileIO.FileSystem.FileExists(sMasterImgPath) Then
                    ClassificationInfo.ScreenShotImgPath = IO.Path.ChangeExtension(IO.Path.GetTempFileName(), ".png")
                    FileIO.FileSystem.CopyFile(sMasterImgPath, ClassificationInfo.ScreenShotImgPath, True)
                    If Not FileIO.FileSystem.FileExists(ClassificationInfo.ScreenShotImgPath) Then
                        'Failed to copy
                        ClassificationInfo.ScreenShotImgPath = ""
                    End If
                End If
            End If

            'Have Clasification Info So Now Clasify
            If Clasify() <> 0 Then
                Return
            End If
            WindowsAPI.MinimiseOntologyStudio()

            'Have Clasified the project now need to create the structure
            CreateItem() 'First create the item in EPDM

            'Now create the xml properties file
            CreateProjectInfoFile()

            'Create The AutoCAD Project Folder
            CreateProjectFolder()

            'Check In the proejct Information File
            CheckInFile()

            'Reference The Item To The Project Information File
            AddItemReference()

            'Check In The Item for this project
            CheckInItem()

            'Move state to WIP to compensate for auto transistion failure on items :- Known SWEPDM bug SPR540720 - Automatic transitions set as first transition do not work for items
            MoveItemtoWIP()
            'check out item for user to work on

            'This should Check Out The Item As Well
            CheckOutFile()

            'The basic Vault structure should now exist

            'Now Create The Actual AutoCAD Project
            ProjectExists = CreateACEProject()

            If Not ProjectExists Then Return

            'List All the drawings In the Project
            ListDrawings()

            'Now need to update the title blocks of all the drawings in the project
            Dim AttributeMaps As Hashtable = MapPDMVariables()
            If AttributeMaps IsNot Nothing Then
                UpdateTitleBlock(AttributeMaps)
            End If

        Catch ex As Exception
            Throw New Exception("Failed Creating New Project", ex)
        End Try


    End Sub

    Private Sub UpdateTitleBlock(ByVal AttributeMap As Hashtable)

    End Sub

    Private Function MapPDMVariables() As Hashtable
        Try
            Dim ManagedXmlDocument As XPathDocument = New XPathDocument(ProjectInfoFile)
            If ManagedXmlDocument Is Nothing Then Throw New Exception(" The Project Info file Could Not Be Opened")
            Dim xmlNavigator As XPathNavigator = ManagedXmlDocument.CreateNavigator()
            If xmlNavigator Is Nothing Then Throw New Exception(" Cannot Navigate The Project Info file")

            'Get the description
            Dim DrawingDescription As String = ""
            Dim xmlNodeItterator As XPathNodeIterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Description']")
            If xmlNodeItterator IsNot Nothing Then
                DrawingDescription = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Revision
            Dim DrawingRevision As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Revision']")
            If xmlNodeItterator IsNot Nothing Then
                DrawingRevision = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Drawn By
            Dim DrawingDesigner As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Drawn By']")
            If xmlNodeItterator IsNot Nothing Then
                DrawingDesigner = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Drawn Date
            Dim DrawingDate As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='DrawnDate']")
            If xmlNodeItterator IsNot Nothing Then
                DrawingDate = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            xmlNodeItterator = Nothing
            xmlNavigator = Nothing
            ManagedXmlDocument = Nothing

            Dim AttributeMap As Hashtable = New Hashtable()
            AttributeMap.Add("ISSUE NUMBER", DrawingRevision)
            AttributeMap.Add("DRAWING NUMBER", ClassificationInfo.Name)
            AttributeMap.Add("DRAWN DATE", DrawingDate)
            AttributeMap.Add("DRAWN BY", DrawingDesigner)
            AttributeMap.Add("DESCRIPTION 1", DrawingDescription)

            Return AttributeMap
        Catch ex As Exception
            Throw New Exception("Failed Mapping PDM Variables", ex)
        End Try

    End Function

    ''' <summary>
    ''' Lists All the Drawings In the Project 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ListDrawings()
        Try
            If Not ProjectExists Then Return
            DrawingsList = New List(Of String)

            Dim rb As ResultBuffer = New ResultBuffer()
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, "c:wd_proj_wdp_data_fnam"))
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, FullProjectName.Replace("\", "/")))
            Dim Res_rb As ResultBuffer = Application.Invoke(rb)
            Dim resultArray As TypedValue() = Res_rb.AsArray()
            Dim DrawingList As TypedValue() = Nothing
            LispNth(resultArray, 4, DrawingList)
            For Each Drawing As TypedValue In DrawingList
                DrawingsList.Add(Drawing.Value.ToString())
            Next
        Catch ex As Exception
            Throw New Exception("Failed Listing Project Drawings", ex)
        End Try

    End Sub

    ''' <summary>
    ''' A .NET equivalent to the Lisp nth function
    ''' </summary>
    ''' <param name="TypedValueList">A List of the paramters to perform the function on</param>
    ''' <param name="Number">the item number of the list</param>
    ''' <param name="Result">the nth element of the list</param>
    ''' <remarks>Will not handle a list inside a list iniside a list</remarks>
    Private Sub LispNth(ByVal TypedValueList As TypedValue(), ByVal Number As Integer, ByRef Result As TypedValue())
        Dim CurrentElement As Integer = 0
        Dim ResultList As List(Of TypedValue) = New List(Of TypedValue)

        For Index As Integer = 0 To TypedValueList.Length - 1
            If CurrentElement = Number Then
                'Return this element
                If TypedValueList(Index).TypeCode = Autodesk.AutoCAD.Runtime.LispDataType.ListBegin Then
                    'this is a list add alements To Whe List
                    Dim InnerListIndex As Integer = Index + 1
                    While TypedValueList(InnerListIndex).TypeCode <> Autodesk.AutoCAD.Runtime.LispDataType.ListEnd
                        ResultList.Add(TypedValueList(InnerListIndex))
                        InnerListIndex += 1
                    End While
                    Exit For
                Else
                    ResultList.Add(TypedValueList(Index))
                    Exit For
                End If
            End If

            If TypedValueList(Index).TypeCode <> Autodesk.AutoCAD.Runtime.LispDataType.ListBegin Then
                CurrentElement += 1
            Else
                'increment until we reach the list end
                CurrentElement += 1
                While TypedValueList(Index).TypeCode <> Autodesk.AutoCAD.Runtime.LispDataType.ListEnd
                    Index += 1
                End While
            End If
        Next
        Result = ResultList.ToArray()
    End Sub
    ''' <summary>
    ''' Copies the Template AutoCAD Project
    ''' </summary>
    ''' <remarks></remarks>
    Private Function CreateACEProject() As Boolean
        Try
            'Get the path of tge template project
            Dim TemplateProject As String = "C:\ProgramData\AutoCAD\AutoCAD2013\AeData\Proj\Ishida Default Project\Ishida Default Project.wdp" 'ItemGatewayDatabase.GetSetting("ManagedDocumentsFolder")

            If Not IO.File.Exists(TemplateProject) Then Return False 'Cannot Find Template Project

            'Construct the path for the new project
            Dim NewProjectName As String = ProjectFolder
            If Not NewProjectName.EndsWith("\") Then NewProjectName &= "\"
            NewProjectName &= ProjectTitle
            NewProjectName &= ".wdp"

            Dim app As AcadApplication = DirectCast(Application.AcadApplication, AcadApplication)
            Dim rb As ResultBuffer = New ResultBuffer()
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, "c:wd_cpyprj_main"))

            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.ListBegin, -1))                             'Start of List
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, TemplateProject.Replace("\", "/")))   'Project To Copy From
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, NewProjectName.Replace("\", "/")))    'New Project
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Nil, -1))                                   'Copy All Drawings
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ProjectFolder.Replace("\", "/")))    'New Project Folder

            'Cant Get This Code to Work will inevestigate later Only a nice to have
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.ListBegin, -1)) 'Start of file flags list
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'Copy wdt file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 1))      'Copy wdl file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wdd file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy _cat.mdb file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy _footprintlookup.mdb
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wd_fam.dat
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wdw file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 1))      'copy loc file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 1))      'copy inst file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wdr file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wdx file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wdi file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Int16, 0))      'copy wdf file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.ListEnd, -1))   'End Of fileFlags Lst

            'Dim wdlFile As String = IO.Path.GetDirectoryName(TemplateProject)
            'If Not wdlFile.EndsWith("\") Then wdlFile &= "\"
            'wdlFile &= IO.Path.GetFileNameWithoutExtension(TemplateProject)

            'Dim locFile As String = wdlFile & ".loc"
            'Dim instFile As String = wdlFile & ".inst"
            'wdlFile &= ".wdl"

            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.ListBegin, -1)) 'Start of file locations list
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdt file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, wdlFile.Replace("\", "/"))) 'wdl file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdd file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      '_cat.mdb file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      '_footprintlookup.mdb
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wd_fam.dat
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdw file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, locFile.Replace("\", "/"))) 'loc file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, instFile.Replace("\", "/"))) 'inst file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdr file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdx file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdi file
            ''rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, ""))      'wdf file
            'rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.ListEnd, -1))   'End Of file locations Lst

            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.ListEnd, -1))   'End Of List

            Dim Res_rb As ResultBuffer = Application.Invoke(rb) 'Create The Project
            If Not IO.File.Exists(NewProjectName) Then Return False 'For some Reason the New Project Was Not Created

            FullProjectName = NewProjectName
            Return True
        Catch ex As Exception
            Throw New Exception("Failed Creating AutoCAD Project", ex)
        End Try

    End Function


    ''' <summary>
    ''' Clasify This Part In Ontology Studio
    ''' </summary>
    ''' <remarks></remarks>
    Private Function Clasify() As Integer
        Try
            Dim ConnectionToOntlogyStudio As OntologyStudioConnection = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.OntologyStudio), OntologyStudioConnection)

            'Create The clasifcation Info Obhject
            Return ConnectionToOntlogyStudio.Clasify(ClassificationInfo)

        Catch ex As Exception
            Throw New Exception("Failed Clasifying Part", ex)
        End Try
    End Function

    ''' <summary>
    ''' Creates An Item for this Part in EPDM, in the correct path obtained by clasifying this part
    ''' </summary>
    ''' <remarks>Use ClassificationInfo and Sets EPDMItemID
    ''' Check in out bug fix added by Tim</remarks>
    Private Sub CreateItem()
        Try
            Dim ConnectionToEPDM As EPDMConnection = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection)
            Dim NewItemPath As String = ConnectionToEPDM.VaultItemRootFolder
            If Not NewItemPath.EndsWith("\") Then NewItemPath &= "\"
            NewItemPath &= ConnectionMgr.EPDM_ConnectionInfo.ManagedDocumentsFolder
            If Not ClassificationInfo.RelativePath.StartsWith("\") Then NewItemPath &= "\"
            NewItemPath &= ClassificationInfo.RelativePath
            EPDMItemID = ConnectionToEPDM.CreateItem(ClassificationInfo.Name, NewItemPath, ClassificationInfo.Description)
        Catch ex As Exception
            Throw New Exception("Failed Creating Item For Part", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Creates The Main Project Information File For Asociating with The Item
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateProjectInfoFile()
        Try
            Dim ProjectInfoFilePath As String = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).VaultRootFolder
            If Not ProjectInfoFilePath.EndsWith("\") Then ProjectInfoFilePath &= "\"
            ProjectInfoFilePath &= ConnectionMgr.EPDM_ConnectionInfo.ManagedDocumentsFolder
            If Not ClassificationInfo.RelativePath.StartsWith("\") Then ProjectInfoFilePath &= "\"
            ProjectInfoFilePath &= ClassificationInfo.RelativePath

            Dim ProjectInfoFileName As String = ClassificationInfo.Name & ".xml"

            ProjectInfoFile = ProjectInfoFilePath
            If Not ProjectInfoFile.EndsWith("\") Then ProjectInfoFile &= "\"
            ProjectInfoFile &= ProjectInfoFileName

            EPDMFileID = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).AddFileToVault(ProjectInfoFilePath, ProjectInfoFileName)
            If EPDMFileID <= 0 Then
                Throw New Exception("Failed Adding Project Information File To EPDM!")
            End If

            Dim InfoFileWriter As New XmlTextWriter(ProjectInfoFile, System.Text.Encoding.UTF8)
            InfoFileWriter.WriteStartDocument(True)
            InfoFileWriter.Formatting = Formatting.Indented
            InfoFileWriter.Indentation = 2

            InfoFileWriter.WriteStartElement("SWEPDM_GENERIC_FILE_PACKAGES")
            InfoFileWriter.WriteStartElement("FILE_PACKAGE")
            InfoFileWriter.WriteAttributeString("TYPE", "ACADE")

            InfoFileWriter.WriteStartElement("DATA_SOURCE")
            ProjectFolder = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).VaultRootFolder
            If Not ProjectFolder.EndsWith("\") And Not ConnectionMgr.EPDM_ConnectionInfo.ManagedDocumentsFolder.StartsWith("\") Then ProjectFolder &= "\"
            ProjectFolder &= ConnectionMgr.EPDM_ConnectionInfo.ManagedDocumentsFolder
            If Not ProjectFolder.EndsWith("\") And Not ClassificationInfo.RelativePath.StartsWith("\") Then ProjectFolder &= "\"
            ProjectFolder &= ClassificationInfo.RelativePath
            If Not ProjectFolder.EndsWith("\") Then ProjectFolder &= "\"
            ProjectFolder &= ClassificationInfo.Name

            InfoFileWriter.WriteAttributeString("PATH", ProjectFolder)
            ProjectTitle = RequestProjectName()
            InfoFileWriter.WriteAttributeString("HEADER_FILE", ProjectTitle & ".wdp")
            InfoFileWriter.WriteEndElement() '</DATA_SOURCE>

            InfoFileWriter.WriteStartElement("VARIABLES")

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Number")
            InfoFileWriter.WriteAttributeString("VALUE", ClassificationInfo.Name)
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Description")
            InfoFileWriter.WriteAttributeString("VALUE", ClassificationInfo.Description)
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Revision")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Drawn By")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "DrawnDate")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Checked By")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Checked Date")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssuedBy")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssuedDate")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "Notes")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteEndElement() '</VARIABLES>

            InfoFileWriter.WriteStartElement("BOM")

            InfoFileWriter.WriteEndElement() '</BOM>

            InfoFileWriter.WriteEndElement() '</FILE_PACKAGE>'
            InfoFileWriter.WriteEndElement() '</SWEPDM_GENERIC_FILE_PACKAGES>'
            InfoFileWriter.WriteEndDocument()
            InfoFileWriter.Close()

        Catch ex As Exception
            Throw New Exception("Failed Adding Project Information File To EPDM!", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Requests The User To Provide a Freindly Short File Name For The AutoCad Project
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestProjectName() As String
        Dim dlgProjectName As New AddinDialogs.ProjectTitleDialog

        ' Configure the dialog box
        dlgProjectName.ShortTitle = ClassificationInfo.Name
        'Show the dialog
        dlgProjectName.ShowDialog()

        If dlgProjectName.ShortTitle.Length <= 0 Or dlgProjectName.ShortTitle.Length > 20 Then
            Return ClassificationInfo.Name
        End If
        If dlgProjectName.ShortTitle.ToLower().Contains(ClassificationInfo.Name.ToLower()) Then
            'The user has already included the itme name in the title
            Return dlgProjectName.ShortTitle
        Else
            Return dlgProjectName.ShortTitle & "[" & ClassificationInfo.Name & "]"
        End If

    End Function

    ''' <summary>
    ''' Checks in the File Represented By this object in the EPDM Vault
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInFile()
        Try
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).CheckInFile(EPDMFileID)
        Catch ex As Exception
            Throw New Exception("Failed checking a file In", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Checks out the File Represented By this object in the EPDM Vault
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckOutFile()
        Try
            Dim CheckOutList As List(Of Integer) = New List(Of Integer)
            CheckOutList.Add(EPDMFileID)
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).CheckOutFiles(CheckOutList)
        Catch ex As Exception
            Throw New Exception("Failed checking out a file", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Checks In The EPDM Item Associated With This object
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInItem()
        Try
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).CheckInFile(EPDMItemID)
        Catch ex As Exception
            Throw New Exception("Failed Checking In The Item", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Checks out The EPDM Item Associated With This object
    ''' </summary>
    ''' <remarks>Tim Addition</remarks>
    Private Sub CheckOutItem()
        Try
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).CheckOutFile(EPDMItemID)
        Catch ex As Exception
            Throw New Exception("Failed Checking out The Item", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Moves an Item to WIP
    ''' </summary>
    ''' <remarks>Tim Addition</remarks>
    Private Sub MoveItemtoWIP()
        Try
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).MoveFileToState(EPDMItemID, "WIP")
        Catch ex As Exception
            Throw New Exception("Failed moving item to WIP state", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Adds A Referencece To the item representing this object in the EPDM Vault to an associate file
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddItemReference()
        Try
            'Note there are no configurations in this type of File 
            Dim ErrStr As String = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).ReferenceFileToItem(EPDMFileID, EPDMItemID, "")
            If Not ErrStr = "S_OK" Then
                Throw New Exception(ErrStr)
            End If
        Catch ex As Exception
            Throw New Exception("Failed To Add Item Reference", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Creates The AutoCad electrcial Project Folder For This is it does not exist 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateProjectFolder()
        Try
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).CreateFolderPath(ProjectFolder)
        Catch ex As Exception
            Throw New Exception("Failed To Create The AutoCAD Electrical Project Folder", ex)
        End Try
    End Sub
End Class
