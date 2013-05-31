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
    Private Const csTitleBlockName As String = "ISHDA STANDARD BORDER"
    Private Const csTitlePageBlockName As String = "TITLEBLOCK"
    Private Const csInfoBlockName As String = "INFOBLOCK"

    Private ClassificationInfo As OntologyStudioData = Nothing
    Private ConnectionMgr As ConnectionManager = Nothing
    Private EPDMItemID As Integer = -1
    Private EPDMFileID As Integer = -1
    Private ProjectInfoFile As String = Nothing
    Private PartNumberPattern As String = "*"
    Private ProjectTitle As String = ""
    Private ProjectFolder As String = ""
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

            'Make sure all files created are added to the vault
            AddAllFilesToEpdm()

            'Now need to update the title blocks of all the drawings in the project
            Dim AttributeMaps As Hashtable = MapPDMVariables()
            If AttributeMaps IsNot Nothing Then
                UpdateTitleBlocks(AttributeMaps)
                UpdateTitlePages(True, True, False)
            End If

            'Ensure that the files have been checked in once
            FirstCheckInACEProjectFiles()



        Catch ex As Exception
            Throw New Exception("Failed Creating New Project", ex)
        End Try


    End Sub

    ''' <summary>
    ''' Just runs A Check In and check out on a new ACE project to ensure all files have been added into the vault 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FirstCheckInACEProjectFiles()
        Try
            Dim EpdmConn As EPDMConnection = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection)
            If EpdmConn IsNot Nothing Then
                Dim FileIDs As List(Of Integer) = New List(Of Integer)
                For Each ProjectDrawing As String In DrawingsList
                    Dim FileID As Integer = EpdmConn.FetchDocumentID(ProjectDrawing)
                    If FileID > 0 Then
                        FileIDs.Add(FileID)
                    End If
                Next
                'Wdp file
                Dim WdpFileID As Integer = EpdmConn.FetchDocumentID(FullProjectName)
                If WdpFileID > -1 Then
                    FileIDs.Add(WdpFileID)
                End If
                'Directory Listing for other project files

                Dim di As New IO.DirectoryInfo(IO.Path.GetDirectoryName(FullProjectName))
                For Each ListedFile As IO.FileInfo In di.GetFiles()
                    If String.Compare(ListedFile.Extension, ".dsd", True) = 0 Or String.Compare(ListedFile.Extension, ".loc", True) = 0 Or String.Compare(ListedFile.Extension, ".inst", True) = 0 Or String.Compare(ListedFile.Extension, ".wdl", True) = 0 Then
                        Dim FileID As Integer = EpdmConn.FetchDocumentID(ListedFile.FullName)
                        If FileID > -1 Then
                            FileIDs.Add(FileID)
                        End If
                    End If
                Next
                'Check In the files
                EpdmConn.CheckInFiles(FileIDs, False)

                'following this need to check that all files are in the correct state if not transition to wip using the no set revision transistion
                For Each FileID As Integer In FileIDs
                    Dim CurrentState As String = ""
                    CurrentState = EpdmConn.QueryFileState(FileID)
                    If CurrentState Is Nothing Then Continue For
                    If String.Compare(CurrentState, "WIP") = 0 Then Continue For 'File is already in wip
                    EpdmConn.MoveFileToState(FileID, "WIP")
                Next
            End If

        Catch ex As Exception
            Throw New Exception("Failed Checking In New Project Files")
        End Try
    End Sub

    ''' <summary>
    ''' Checks If All Valid files are added to the vault for this project
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddAllFilesToEpdm()
        Dim EpdmConn As EPDMConnection = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection)
        If EpdmConn IsNot Nothing Then
            For Each ProjectDrawing As String In DrawingsList
                Dim FileID As Integer = EpdmConn.FetchDocumentID(ProjectDrawing)
                If FileID = -1 Then
                    'Not added So add it
                    Try
                        EpdmConn.AddFileToVault(ProjectDrawing, IO.Path.GetDirectoryName(ProjectDrawing), "")
                    Catch
                    End Try
                End If

            Next
            'Test Wdp file
            Dim WdpFileID As Integer = EpdmConn.FetchDocumentID(FullProjectName)
            If WdpFileID = -1 Then
                Try
                    EpdmConn.AddFileToVault(FullProjectName, IO.Path.GetDirectoryName(FullProjectName), "")
                Catch
                End Try
            End If
            'Directory Listing

            Dim di As New IO.DirectoryInfo(IO.Path.GetDirectoryName(FullProjectName))
            For Each ListedFile As IO.FileInfo In di.GetFiles()
                If String.Compare(ListedFile.Extension, ".dsd", True) = 0 Or String.Compare(ListedFile.Extension, ".loc", True) = 0 Or String.Compare(ListedFile.Extension, ".inst", True) = 0 Or String.Compare(ListedFile.Extension, ".wdl", True) = 0 Then
                    Dim FileID As Integer = EpdmConn.FetchDocumentID(ListedFile.FullName)
                    If FileID = -1 Then
                        'Not added So add it
                        Try
                            EpdmConn.AddFileToVault(ListedFile.FullName, IO.Path.GetDirectoryName(FullProjectName), "")
                        Catch
                        End Try
                    End If

                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Updates the Title Page Information Blocks On The project
    ''' </summary>
    ''' <param name="NewProject"> Set To True If this is a new project, this causes the first time parameters to be set</param>
    ''' <param name="UpdatedProject">Set To True To Update The Project Dates And time</param>
    ''' <remarks></remarks>
    Private Sub UpdateTitlePages(Optional ByVal NewProject As Boolean = True, Optional ByVal UpdatedProject As Boolean = True, Optional ByVal SetCADData As Boolean = False)
        Dim AttributeMap As Hashtable = New Hashtable()

        If SetCADData Then
            'Prompt The User To Provide Deatails
            'Set values from Dialog
            AttributeMap.Add("PWR", "")
            AttributeMap.Add("DESCRIPTION", "")
            AttributeMap.Add("CUSTOMER", "")
            AttributeMap.Add("FREQ", "")
            AttributeMap.Add("CSA", "")
            AttributeMap.Add("AVE_LOAD", "")
            AttributeMap.Add("PEAK_LOAD", "")
            AttributeMap.Add("PHASE", "")
            AttributeMap.Add("VOLTAGE", "")
            AttributeMap.Add("RATING", "")
        End If

        If NewProject Then
            'We are not running this on a new project
            'Is a new project so just populate the non editable fields
            AttributeMap.Add("CREATED_BY", DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).LoggedInUser.ToUpper())
            AttributeMap.Add("CREATED_ON", System.DateTime.Now.ToShortDateString())
            AttributeMap.Add("DRAWING_NUM", ClassificationInfo.Name)
            AttributeMap.Add("PROJECT_NAME", ClassificationInfo.Description)
        End If

        If UpdatedProject Then 'The project has been updated so set update details
            AttributeMap.Add("PAGE_COUNT", GetSheetMax().ToString())
            AttributeMap.Add("MOD_BY", DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).LoggedInUser.ToUpper())
            AttributeMap.Add("LAST_MOD", System.DateTime.Now.ToShortDateString())
            AttributeMap.Add("PROJECT_NUMBER", GetProjectNumber())
        End If

        'The title pages should always be on the GD01 and GD02 Drawings
        Try
            For Each ProjectDrawing As String In DrawingsList
                If ProjectDrawing.ToUpper().Contains("GD01") Or ProjectDrawing.ToUpper().Contains("GD02") Then
                    Dim db As Database = New Database(False, True)

                    If db Is Nothing Then Throw New Exception("Failed to obtain Drawing Database")
                    Using db
                        db.ReadDwgFile(ProjectDrawing, IO.FileShare.ReadWrite, False, "")
                        UpdateAttributesInDatabase(db, csTitlePageBlockName, AttributeMap)
                        UpdateAttributesInDatabase(db, csInfoBlockName, AttributeMap)
                        db.SaveAs(ProjectDrawing, DwgVersion.Newest)
                    End Using
                End If

            Next
        Catch ex As Exception
            Throw New Exception("Error Updating Title Page Blocks in Project")
        End Try
    End Sub

    ''' <summary>
    ''' Gets The Sheet Count For this Project
    ''' </summary>
    ''' <returns>The sheet Count For the Project</returns>
    ''' <remarks></remarks>
    Private Function GetSheetMax() As Integer
        Try
            If Not ProjectExists Then Return -1

            Dim rb As ResultBuffer = New ResultBuffer()
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, "c:wd_proj_wdp_data_fnam"))
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, FullProjectName.Replace("\", "/")))
            Dim Res_rb As ResultBuffer = Application.Invoke(rb)
            Dim resultArray As TypedValue() = Res_rb.AsArray()
            Dim Count As Integer = -1
            Integer.TryParse(resultArray(0).Value.ToString(), Count)
            Return Count
        Catch ex As Exception
            Throw New Exception("Failed Obtaining Drawing Count", ex)
        End Try
    End Function

    ''' <summary>
    ''' Gets the Project Number for the project, this is the only title set with the Autocad project titles
    ''' </summary>
    ''' <returns>Project title or empty string</returns>
    ''' <remarks>Not tested</remarks>
    Private Function GetProjectNumber() As String
        Try
            If Not ProjectExists Then Return ""

            Dim rb As ResultBuffer = New ResultBuffer()
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, "c:wd_proj_wdp_data_fnam"))
            rb.Add(New TypedValue(Autodesk.AutoCAD.Runtime.LispDataType.Text, FullProjectName.Replace("\", "/")))
            Dim Res_rb As ResultBuffer = Application.Invoke(rb)
            Dim resultArray As TypedValue() = Res_rb.AsArray()
            Dim TitlesList As TypedValue() = Nothing
            LispNth(resultArray, 1, TitlesList)
            If TitlesList IsNot Nothing Then
                If TitlesList.Length >= 2 Then
                    Return TitlesList(1).Value.ToString()
                End If
            End If
            Return ""
        Catch ex As Exception
            Throw New Exception("Failed Obtaining Drawing Count", ex)
        End Try
    End Function

    ''' <summary>
    ''' Updates The Basic Title Block Information On All Valid Drawings in the project
    ''' </summary>
    ''' <param name="AttributeMap">A Mapped Table of Attributes and Values</param>
    ''' <remarks>Should Only be used when creating a new project, as this function sets the basic title block information</remarks>
    Private Sub UpdateTitleBlocks(ByVal AttributeMap As Hashtable)
        Try
            For Each ProjectDrawing As String In DrawingsList
                Dim db As Database = New Database(False, True)

                If db Is Nothing Then Throw New Exception("Failed to obtain Drawing Database")
                Using db
                    db.ReadDwgFile(ProjectDrawing, IO.FileShare.ReadWrite, False, "")
                    UpdateAttributesInDatabase(db, csTitleBlockName, AttributeMap)
                    db.SaveAs(ProjectDrawing, DwgVersion.Newest)
                End Using
            Next
        Catch ex As Exception
            Throw New Exception("Error Updating Title Blocks in Project")
        End Try
    End Sub

    ''' <summary>
    ''' Reads the base variables from Enterprise EPDM syetm and maps them to Autocad Drawings TitleBlock Attribute names 
    ''' </summary>
    ''' <returns>A Hashmap, Keys are the Autocad Attribute Name, Values are valeus to be written to the attribute</returns>
    ''' <remarks></remarks>
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
                xmlNodeItterator.MoveNext()
                DrawingDescription = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Revision
            Dim DrawingRevision As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Revision']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                DrawingRevision = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Drawn By
            Dim DrawingDesigner As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Drawn By']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                DrawingDesigner = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Drawn Date
            Dim DrawingDate As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='DrawnDate']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                DrawingDate = xmlNodeItterator.Current.GetAttribute("VALUE", "")
                If DrawingDate.Length > 0 Then
                    Dim DateParts As String() = DrawingDate.Split("/"c)
                    If DateParts(2).Length > 2 Then
                        DrawingDate = DateParts(0) & "/" & DateParts(1) & "/" & DateParts(2).Substring(2)
                    End If
                End If
            End If

            'Get Drawing Number
            Dim DrawingNumber As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Number']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                DrawingNumber = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Checked By Name
            Dim CheckedByName As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Checked By']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                CheckedByName = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get CheckedDate
            Dim CheckedDate As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='Checked Date']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                CheckedDate = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Issued By Name
            Dim IssuedByName As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssuedBy']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssuedByName = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Issued Date
            Dim IssuedDate As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssuedDate']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssuedDate = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Issue History 1 Number
            Dim IssueHist1Number As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory1Number']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist1Number = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            'Get Issue History 1 Name
            Dim IssueHist1Name As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory1Name']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist1Name = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            'Get Issue History 1 Description
            Dim IssueHist1Desc As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory1Desc']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist1Desc = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            'Get Issue History 1 Date
            Dim IssueHist1Date As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory1Date']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist1Number = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            'Get Issue History 2 Number
            Dim IssueHist2Number As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory2Number']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist2Number = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            'Get Issue History 2 Name
            Dim IssueHist2Name As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory2Name']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist2Name = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            'Get Issue History 2 Description
            Dim IssueHist2Desc As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory2Desc']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist2Desc = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If
            'Get Issue History 1 Date
            Dim IssueHist2Date As String = ""
            xmlNodeItterator = xmlNavigator.Select("SWEPDM_GENERIC_FILE_PACKAGES/FILE_PACKAGE[@TYPE='ACADE']/VARIABLES/VARIABLE[@NAME='IssueHistory2Date']")
            If xmlNodeItterator IsNot Nothing Then
                xmlNodeItterator.MoveNext()
                IssueHist2Number = xmlNodeItterator.Current.GetAttribute("VALUE", "")
            End If

            xmlNodeItterator = Nothing
            xmlNavigator = Nothing
            ManagedXmlDocument = Nothing

            Dim AttributeMap As Hashtable = New Hashtable()
            AttributeMap.Add("ISSUE_NUM", DrawingRevision)
            AttributeMap.Add("DRAWING_NO", DrawingNumber)
            AttributeMap.Add("DRWN_DAT", DrawingDate)
            AttributeMap.Add("DRWN_BY", DrawingDesigner)
            AttributeMap.Add("DESCRIPTION_1", DrawingDescription)
            AttributeMap.Add("CHECK_BY", CheckedByName)
            AttributeMap.Add("CHECK_DAT", CheckedDate)
            AttributeMap.Add("ISSUE_BY", IssuedByName)
            AttributeMap.Add("ISS_DATE", IssuedDate)
            AttributeMap.Add("IS1", IssueHist1Number)
            AttributeMap.Add("IS_BY1", IssueHist1Name)
            AttributeMap.Add("ISS_DESC_1", IssueHist1Desc)
            AttributeMap.Add("IS_DATE_1", IssueHist2Date)
            AttributeMap.Add("IS2", IssueHist2Number)
            AttributeMap.Add("IS_BY2", IssueHist2Name)
            AttributeMap.Add("ISS_DESC_2", IssueHist2Desc)
            AttributeMap.Add("IS_DATE2", IssueHist2Date)
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
                If IO.File.Exists(Drawing.Value.ToString().Replace("/", "\")) Then
                    DrawingsList.Add(Drawing.Value.ToString().Replace("/", "\"))
                End If
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
    ''' <remarks>Will not handle a list inside a list iniside a list without modification</remarks>
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
            Dim TemplateProject As String = ItemGatewayDatabase.GetSetting("AceTemplateProject")

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

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory1Number")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory1Name")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory1Desc")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory1Date")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory2Number")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory2Name")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory2Desc")
            InfoFileWriter.WriteAttributeString("VALUE", "")
            InfoFileWriter.WriteEndElement() '</VARIABLE>'

            InfoFileWriter.WriteStartElement("VARIABLE")
            InfoFileWriter.WriteAttributeString("NAME", "IssueHistory2Date")
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


    ''' <summary>
    ''' Updates the given attributes in all blocks on the activedrawing
    ''' </summary>
    ''' <param name="BlockName"></param>
    ''' <param name="AttbMap"></param>
    ''' <remarks></remarks>
    Private Sub UpdateAttribute(ByVal BlockName As String, ByVal AttbMap As Hashtable)
        Try
            Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
            UpdateAttributesInDatabase(db, BlockName, AttbMap)
        Catch ex As Exception
            Throw New Exception("Error Updating Attributes")
        End Try
    End Sub

    ''' <summary>
    ''' Updates The given attributes of every block in an autocad Drawings Database
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="BlockName"></param>
    ''' <param name="AttbMap"></param>
    ''' <remarks></remarks>
    Private Sub UpdateAttributesInDatabase(ByVal db As Database, ByVal BlockName As String, ByVal AttbMap As Hashtable)
        Try

            Dim msID As ObjectId = Nothing
            Dim psID As ObjectId = Nothing

            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                msID = bt(BlockTableRecord.ModelSpace)
                psID = bt(BlockTableRecord.PaperSpace)
                tr.Commit()
            End Using

            Dim msCount As Integer = UpdateAttributesInBlock(db, msID, AttbMap, BlockName)
            Dim psCount As Integer = UpdateAttributesInBlock(db, psID, AttbMap, BlockName)
        Catch ex As Exception
            Throw New Exception("Error Updating Attributes In Database")
        End Try
    End Sub

    ''' <summary>
    ''' Updates a list of attributes in a block
    ''' </summary>
    ''' <param name="btrId">The Plock Table record for either paperspace of model space</param>
    ''' <param name="attbMap">A Map of attributes and Values</param>
    ''' <param name="blockName">The name of the block</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdateAttributesInBlock(ByVal db As Database, ByVal btrId As ObjectId, ByVal attbMap As Hashtable, ByVal blockName As String) As Integer
        Try
            Dim changedCount As Integer = 0

            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = CType(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)

                For Each entId As ObjectId In btr
                    Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                    If ent IsNot Nothing Then
                        Dim br As BlockReference = TryCast(ent, BlockReference)
                        If br IsNot Nothing Then
                            Dim bd As BlockTableRecord = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                            'Check ownerID for layout info
                            If String.Compare(bd.Name, blockName, True) = 0 Then
                                For Each arId As ObjectId In br.AttributeCollection
                                    Dim obj As DBObject = tr.GetObject(arId, OpenMode.ForRead)
                                    Dim ar As AttributeReference = TryCast(obj, AttributeReference)
                                    If ar IsNot Nothing Then
                                        For Each attbName As String In attbMap.Keys
                                            If String.Compare(ar.Tag, attbName, True) = 0 Then
                                                ar.UpgradeOpen()
                                                ar.TextString = attbMap(attbName).ToString()
                                                Dim wdb As Database = HostApplicationServices.WorkingDatabase
                                                HostApplicationServices.WorkingDatabase = db
                                                ar.AdjustAlignment(db)
                                                HostApplicationServices.WorkingDatabase = wdb
                                                ar.DowngradeOpen()
                                                changedCount += 1
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                            changedCount += UpdateAttributesInBlock(db, br.BlockTableRecord, attbMap, blockName)
                        End If
                    End If
                Next
                tr.Commit()
            End Using
            Return changedCount
        Catch ex As Exception
            Throw New Exception("Error Updating Attributes In Block")
        End Try
    End Function
End Class
