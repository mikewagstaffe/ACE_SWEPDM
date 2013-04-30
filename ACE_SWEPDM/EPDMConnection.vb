Imports EdmLib
Imports System.Collections.Generic



Public Class EPDMConnection
    Private edmVault As IEdmVault12 = Nothing
    Private edmCurrentUser As IEdmUser10 = Nothing

    Public Shared NumberVariableName As String = "Number"
    Public Shared DescriptionVariableName As String = "Description"
    Public Shared RevisionVariableName As String = "Revision"

    ''' <summary>
    ''' Make A New connection Into The EPDM Vault
    ''' </summary>
    ''' <param name="VaultName">Name Of The Local Vault</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal VaultName As String)
        edmVault = DirectCast(New EdmVault5(), IEdmVault12)
        edmVault.LoginAuto(VaultName, 0)
        Dim UserID As Integer = edmVault.GetLoggedInWindowsUserID(edmVault.Name)
        edmCurrentUser = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_User, UserID), IEdmUser10)
    End Sub

    ''' <summary>
    ''' Gets The Document ID of a file at the gicen location
    ''' </summary>
    ''' <param name="Path">Path where he file is located</param>
    ''' <returns>Item ID or -1 if not found</returns>
    ''' <remarks></remarks>
    Public Function FetchDocumentID(ByVal Path As String) As Integer
        Try
            Dim objFile As Object = edmVault.GetFileFromPath(Path)
            If objFile Is Nothing Then Return -1
            Dim edmFileId As Integer = DirectCast(objFile, IEdmFile8).ID
            objFile = Nothing
            Return edmFileId
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetLatestVersionNumber(ByVal DocumentID As Integer) As Integer
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, DocumentID), IEdmFile8)
            If edmFile Is Nothing Then Return -1
            Return edmFile.CurrentVersion
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Checks if document is locked
    ''' </summary>
    ''' <param name="DocumentId"></param>
    ''' <param name="CurrentUser">False if not bothered which user</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Islocked(ByVal DocumentId As Integer, Optional ByVal CurrentUser As Boolean = True) As Boolean
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, DocumentId), IEdmFile8)
            If edmFile Is Nothing Then Throw New Exception("Failed To Get File From EPDM")
            If Not edmFile.IsLocked And Not CurrentUser Then
                Return True
            ElseIf edmFile.IsLocked And CurrentUser Then
                If edmFile.LockedByUser.ID = edmCurrentUser.ID Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Who has a file locked 
    ''' </summary>
    ''' <param name="DocumentId"></param>
    ''' <returns> A string containing the users full name and computer</returns>
    ''' <remarks></remarks>
    Public Function IsLockedBy(ByVal DocumentId As Integer) As String
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, DocumentId), IEdmFile8)
            If edmFile Is Nothing Then Throw New Exception("Failed To Get File From EPDM")
            If edmFile.IsLocked Then
                Dim edmUser As IEdmUser9 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_User, edmFile.LockedByUser.ID), IEdmUser9)
                Return edmUser.FullName & " on computer " & edmFile.LockedOnComputer
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Creates A New Item in the EPDM Vault, with the given name in the given path. If the Path does not exist, 
    ''' the path is first created
    ''' </summary>
    ''' <param name="Name">Name for new item</param>
    ''' <param name="Path">Folder Path for the new item</param>
    ''' <param name="Description"></param>
    ''' <returns>The ID of the new Item</returns>
    ''' <remarks></remarks>
    Public Function CreateItem(ByVal Name As String, ByVal Path As String, ByVal Description As String) As Integer
        'The interface to create an item
        Dim edmBatchItemGen As IEdmBatchItemGeneration2 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchItemGeneration), IEdmBatchItemGeneration2)
        'Construct the correct Path for the item
        Dim ItemFolderPath As String = ""
        If Not Path.ToUpper().StartsWith(edmVault.ItemRootFolder.LocalPath.ToUpper) Then
            ItemFolderPath = edmVault.ItemRootFolder.LocalPath
            If Not ItemFolderPath.EndsWith("\") And Not Path.StartsWith("\") Then ItemFolderPath &= "\"
        End If
        ItemFolderPath &= Path

        If edmVault.GetFolderFromPath(ItemFolderPath) Is Nothing Then
            'we need to create it
            CreateFolderPath(Path)
            If edmVault.GetFolderFromPath(ItemFolderPath) Is Nothing Then Throw New Exception("Failed To Create Item Folder")
        End If
        Dim edmItemFolder As IEdmFolder6 = DirectCast(edmVault.GetFolderFromPath(ItemFolderPath), IEdmFolder6)
        edmBatchItemGen.AddSelection2(DirectCast(edmVault, EdmVault5), Nothing, Name, edmItemFolder.ID)
        edmItemFolder = Nothing
        edmBatchItemGen.CreateTree(0, EdmItemGenerationFlags.Eigcf_Nothing)
        Dim ppoItems As Array = Nothing
        Dim pbOpen As Boolean = False
        edmBatchItemGen.GenerateItems(0, ppoItems, False)
        edmBatchItemGen = Nothing
        If ppoItems.Length <= 0 Then Throw New Exception("Failed To Create A New Item")
        Dim ItemId As Integer = DirectCast(ppoItems.GetValue(0), EdmGenItemInfo).mlItemID
        Dim edmVarMgr As IEdmVariableMgr7 = CType(edmVault, IEdmVariableMgr7)
        Dim edmBatchVarUpdate As IEdmBatchUpdate2 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchUpdate), IEdmBatchUpdate2)
        Dim edmVar As IEdmVariable5 = Nothing
        'Set Part number variable
        edmVar = edmVarMgr.GetVariable(DirectCast(EPDMConnection.NumberVariableName, Object))
        edmBatchVarUpdate.SetVar(ItemId, edmVar.ID, DirectCast(Name, Object), "", EdmBatchFlags.EdmBatch_AllConfigs)
        'Set Description variable
        edmVar = edmVarMgr.GetVariable(DirectCast(EPDMConnection.DescriptionVariableName, Object))
        edmBatchVarUpdate.SetVar(ItemId, edmVar.ID, DirectCast(Description, Object), "", EdmBatchFlags.EdmBatch_AllConfigs)
        Dim poRetErrors As Array = Nothing
        edmBatchVarUpdate.CommitUpdate(poRetErrors)
        edmVarMgr = Nothing
        edmVar = Nothing
        edmBatchVarUpdate = Nothing
        Return ItemId
    End Function

    ''' <summary>
    ''' Creates A New Folder Path in the EPDM Vault
    ''' </summary>
    ''' <param name="Path">The full folder path to create</param>
    ''' <remarks>all non existent folders of the path are created</remarks>
    ''' <revision number="1">M.Wagstaffe (25/04/13) - Added Check At Begining Of Routine To Confirm Folder Path Does Not Exist </revision>
    Public Sub CreateFolderPath(ByVal Path As String)
        Try
            ''''Revision 1
            Dim edmFolder As IEdmFolder6 = DirectCast(edmVault.GetFolderFromPath(Path), IEdmFolder6)
            If edmFolder IsNot Nothing Then
                'The folder Already Exists
                Return
            End If
            '''''End Revision 1

            If Path.ToUpper().StartsWith(edmVault.ItemRootFolder.LocalPath.ToUpper()) Then
                'Creating An Item Folder
                Dim PathElements As String() = Path.Split("\"c)
                Dim CreatedFolder As IEdmFolder6 = edmVault.ItemRootFolder
                For Index As Integer = 3 To PathElements.Length - 1 ' ignore the first 2 C: and item folder always exist
                    If PathElements(Index) Is Nothing Then
                        Continue For
                    End If
                    If PathElements(Index).Length <= 0 Then
                        Continue For
                    End If
                    If CreatedFolder Is Nothing Then
                        Exit For
                    End If
                    Dim edmExistFolder As IEdmFolder6 = Nothing
                    Try
                        edmExistFolder = DirectCast(CreatedFolder.GetSubFolder(PathElements(Index)), IEdmFolder6)
                    Catch
                        edmExistFolder = Nothing
                    End Try
                    If edmExistFolder Is Nothing Then
                        CreatedFolder = DirectCast(CreatedFolder.AddFolder(0, PathElements(Index)), IEdmFolder6)
                    Else
                        CreatedFolder = edmExistFolder
                    End If
                Next

                If CreatedFolder Is Nothing Then
                    Throw New Exception("Failure In Creation Of Item Folder")
                End If
            ElseIf Path.ToUpper().StartsWith(edmVault.RootFolderPath.ToUpper()) Then
                'Creating A Main Folder
                Dim BatchFolderCreate As IEdmBatchAddFolders = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchAddFolders), IEdmBatchAddFolders)

                BatchFolderCreate.AddFolder(edmVault.RootFolderID, Path.ToUpper().Replace(edmVault.RootFolderPath.ToUpper(), ""), 0)
                Dim CreateReturn As Array = Nothing
                BatchFolderCreate.Create(0, CreateReturn)
            Else
                'Invalid Path
                Throw New Exception("Creating Of Folder Failed: Invalid Folder Location")
            End If

        Catch ex As Exception
            Throw New Exception("Failed Creating Folder in Vault", ex)
        End Try
    End Sub
    ''' <summary>
    ''' Reads a list of variable values from Enterprise PDM
    ''' </summary>
    ''' <param name="FileID">The ID of the file to read the variables from </param>
    ''' <param name="ConfigurationName">Configuration of the file to read variables from "" = all configs</param>
    ''' <param name="MappedVariablesList">List a variables to read</param>
    ''' <remarks>Also gets Attribute Data For SLDDRW files, may need to revist</remarks>
    Public Sub ReadMappedVariables(ByVal FileID As Integer, ByVal ConfigurationName As String, ByRef MappedVariablesList As List(Of MappedVariable))
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            If edmFile Is Nothing Then
                Throw New Exception("Failed To Get File Interface")
            End If
            Dim edmVariableEnum As IEdmEnumeratorVariable8 = DirectCast(edmFile.GetEnumeratorVariable(), IEdmEnumeratorVariable8)
            If edmVariableEnum Is Nothing Then
                Throw New Exception("Failed to Get Variable Enumerator Interface")
            End If
            For Index As Integer = 0 To MappedVariablesList.Count - 1
                Dim VariableValue As Object = Nothing
                edmVariableEnum.GetVar(MappedVariablesList(Index).EPDMName, ConfigurationName, VariableValue)
                MappedVariablesList(Index).WMValue = DirectCast(VariableValue, String)
                Dim edmVarMgr As IEdmVariableMgr7 = CType(edmVault, IEdmVariableMgr7)
                Dim edmVar As IEdmVariable5 = edmVarMgr.GetVariable(DirectCast(MappedVariablesList(Index).EPDMName, Object))
                If edmVar Is Nothing Then Continue For
                MappedVariablesList(Index).EPDMType = edmVar.VariableType
                Dim pos As IEdmPos5 = edmVar.GetFirstAttributePosition("SLDDRW")
                If pos Is Nothing Then Continue For
                Dim att As IEdmAttribute5 = edmVar.GetNextAttribute(pos)
                If att Is Nothing Then Continue For
                MappedVariablesList(Index).SlddrwName = att.Name
                pos = Nothing
                att = Nothing
                edmVar = Nothing
                edmVarMgr = Nothing
            Next
            edmVariableEnum.CloseFile(False)
            edmVariableEnum = Nothing
            Exit Sub
        Catch ex As Exception
            Throw New Exception("Failed Reading File Variables", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Sets All Variables in Enterprise PDM with 
    ''' </summary>
    ''' <param name="FileID">ID of the File TO Write The variables to</param>
    ''' <param name="ConfigurationName">Configuration of the file to write variables to "" = all configs</param>
    ''' <param name="MappedVariablesList">List a variables to write to</param>
    ''' <remarks>Updates The argument MappedVariablesList with EPDM values</remarks>
    Public Sub WriteMappedVariables(ByVal FileID As Integer, ByVal ConfigurationName As String, ByRef MappedVariablesList As List(Of MappedVariable))
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)

            If edmFile Is Nothing Then Throw New Exception("Failed To Get File From EPDM")
            If Not edmFile.IsLocked Or edmFile.LockedByUser.ID <> edmCurrentUser.ID Then Throw New Exception("No Permission To Edit File")
            Dim edmBatchVarUpdate As IEdmBatchUpdate2 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchUpdate), IEdmBatchUpdate2)
            Dim edmVarMgr As IEdmVariableMgr7 = CType(edmVault, IEdmVariableMgr7)

            For Index As Integer = 0 To MappedVariablesList.Count - 1
                Dim edmVar As IEdmVariable5 = edmVarMgr.GetVariable(DirectCast(MappedVariablesList(Index).EPDMName, Object))
                If edmVar Is Nothing Then Throw New Exception("Failed To Find a EPDM Variable")
                MappedVariablesList(Index).EPDMType = edmVar.VariableType
                Select Case MappedVariablesList(Index).EPDMType
                    Case EdmVariableType.EdmVarType_None
                        Continue For
                    Case EdmVariableType.EdmVarType_Bool
                        Dim VarValue As Boolean = False
                        If String.Compare(MappedVariablesList(Index).WMValue, MappedVariablesList(Index).TrueValue, True) = 0 Then
                            VarValue = True
                        End If
                        MappedVariablesList(Index).EPDMValue = VarValue
                    Case EdmVariableType.EdmVarType_Date
                        If MappedVariablesList(Index).WMValue.Length <= 0 Then Continue For
                        Dim VarValue As DateTime = Nothing
                        If Not TryParseWmDate(MappedVariablesList(Index).WMValue, MappedVariablesList(Index).DateFormat, VarValue) Then Continue For
                        MappedVariablesList(Index).EPDMValue = VarValue
                    Case EdmVariableType.EdmVarType_Float
                        If MappedVariablesList(Index).WMValue.Length <= 0 Then Continue For
                        Dim VarValue As Single = 0.0
                        If Not Single.TryParse(MappedVariablesList(Index).WMValue, VarValue) Then Continue For
                        MappedVariablesList(Index).EPDMValue = VarValue
                    Case EdmVariableType.EdmVarType_Int
                        If MappedVariablesList(Index).WMValue.Length <= 0 Then Continue For
                        Dim VarValue As Integer = 0
                        If Not Integer.TryParse(MappedVariablesList(Index).WMValue, VarValue) Then Continue For
                        MappedVariablesList(Index).EPDMValue = VarValue
                    Case EdmVariableType.EdmVarType_Text
                        If MappedVariablesList(Index).WMValue.Length <= 0 Then Continue For
                        MappedVariablesList(Index).EPDMValue = MappedVariablesList(Index).WMValue
                    Case Else
                        Continue For
                End Select

                If ConfigurationName.Length = 0 Then
                    edmBatchVarUpdate.SetVar(edmFile.ID, edmVar.ID, MappedVariablesList(Index).EPDMValue, "", EdmBatchFlags.EdmBatch_AllConfigs)
                Else
                    edmBatchVarUpdate.SetVar(edmFile.ID, edmVar.ID, MappedVariablesList(Index).EPDMValue, ConfigurationName, EdmBatchFlags.EdmBatch_Nothing)
                End If
                edmVar = Nothing
            Next
            'Set The AutoWip Variable
            Dim edmAutoWIPVar As IEdmVariable5 = edmVarMgr.GetVariable(DirectCast("AutoWIP", Object))
            Dim AutoWipValue As Object = "1"
            edmBatchVarUpdate.SetVar(edmFile.ID, edmAutoWIPVar.ID, AutoWipValue, "", EdmBatchFlags.EdmBatch_AllConfigs)
            Dim poRetErrors As Array = Nothing
            If edmBatchVarUpdate.CommitUpdate(poRetErrors) <> 0 Then
                Throw New Exception("Failed Setting Variable Values In EPDM")
            End If

        Catch ex As Exception
            Throw New Exception("Failed Setting Variables In EPDM", ex)
        End Try
    End Sub


    ''' <summary>
    ''' Adds A Given File to the EPDM Vault Storing it in the given path. The path is created if it does not exist
    ''' </summary>
    ''' <param name="LocalFileName">Path of file to Add</param>
    ''' <param name="VaultPath">Location In the vault to store the file</param>
    ''' <param name="NewFileName"></param>
    ''' <returns>The ID of the added file</returns>
    ''' <remarks></remarks>
    Public Function AddFileToVault(ByVal LocalFileName As String, ByVal VaultPath As String, Optional ByVal NewFileName As String = "") As Integer
        Try
            Dim edmBatchAdd As IEdmBatchAdd2 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchAdd), IEdmBatchAdd2)
            Dim edmFolder As IEdmFolder6 = DirectCast(edmVault.GetFolderFromPath(VaultPath), IEdmFolder6)
            If edmFolder Is Nothing Then
                'the folder does not exist so create it
                CreateFolderPath(VaultPath)
                edmFolder = DirectCast(edmVault.GetFolderFromPath(VaultPath), IEdmFolder6)
            End If
            If edmFolder Is Nothing Then Throw New Exception("Failed To Find Destination Folder")
            edmBatchAdd.AddFileFromPath(LocalFileName, edmFolder.ID, 0, NewFileName, EdmAddFlag.EdmAdd_Simple + EdmAddFlag.EdmAdd_Refresh)
            Dim ppoRetFiles As Array = Nothing

            If edmBatchAdd.CommitAdd(0, ppoRetFiles, EdmBatchAddFlag.EdmBaf_Nothing) <> 1 Then
                edmBatchAdd = Nothing
                Throw New Exception("Failed To Add File: " & LocalFileName & " To EPDM")
            End If
            edmBatchAdd = Nothing
            If DirectCast(ppoRetFiles.GetValue(0), EdmFileInfo).mhResult <> 0 Then Throw New Exception("Failed To Add File: " & LocalFileName & " To EPDM")

            Return DirectCast(ppoRetFiles.GetValue(0), EdmFileInfo).mlFileID
        Catch ex As Exception
            Throw New Exception("Failed Adding File To Vault", ex)
        End Try
    End Function

    ''' <summary>
    ''' Adds A Given File to the EPDM Vault Storing it in the given path. The path is created if it does not exist
    ''' </summary>
    ''' <param name="VaultPath">Location In the vault to store the file</param>
    ''' <param name="NewFileName">The name To use for the file</param>
    ''' <returns>The ID of the added file</returns>
    ''' <remarks></remarks>
    Public Function AddFileToVault(ByVal VaultPath As String, ByVal NewFileName As String) As Integer
        Try
            Dim edmFolder As IEdmFolder6 = DirectCast(edmVault.GetFolderFromPath(VaultPath), IEdmFolder6)
            If edmFolder Is Nothing Then
                'the folder does not exist so create it
                CreateFolderPath(VaultPath)
                edmFolder = DirectCast(edmVault.GetFolderFromPath(VaultPath), IEdmFolder6)
            End If
            If edmFolder Is Nothing Then Throw New Exception("Failed To Find Destination Folder")
            Dim FileNameToAdd As String = VaultPath
            If Not FileNameToAdd.EndsWith("\") Then
                FileNameToAdd &= "\"
            End If
            FileNameToAdd &= NewFileName
            Return edmFolder.AddFile(0, "", FileNameToAdd, EdmAddFlag.EdmAdd_Simple + EdmAddFlag.EdmAdd_Refresh)

        Catch ex As Exception
            Throw New Exception("Failed Adding File To Vault", ex)
        End Try
    End Function
    ''' <summary>
    ''' Gets The Path of A File in the EPDM Vault given the File ID
    ''' </summary>
    ''' <param name="FileID">ID of the file to retrieve the path for</param>
    ''' <returns>The full path to the file in the vault</returns>
    ''' <remarks></remarks>
    Public Function GetFilePath(ByVal FileID As Integer) As String
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            If edmFile Is Nothing Then Throw New Exception("Failed To Get File Interface")
            Return edmFile.GetLocalPath(edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID)
        Catch ex As Exception
            Throw New Exception("Failed Adding File To Vault", ex)
        End Try
    End Function

    ''' <summary>
    ''' Gets The latest version of a file
    ''' </summary>
    ''' <param name="FileID">ID of the File in EPDM</param>
    ''' <returns>The Latest Version number</returns>
    ''' <remarks></remarks>
    Public Function GetLastestVersionNo(ByVal FileID As Integer) As Integer
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            If edmFile Is Nothing Then Throw New Exception("Failed To Get File Interface")
            Dim Version As Integer = edmFile.CurrentVersion
            edmFile = Nothing
            Return Version
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Searches The EPDM Vault for a part and returns the path to the id of the item
    ''' </summary>
    ''' <param name="PartNumber">Part number To Find</param>
    ''' <returns>Item ID of found item -1 if not found</returns>
    ''' <remarks></remarks>
    Public Function FindItem(ByVal PartNumber As String) As Integer
        Try

            Dim edmFile As IEdmFile8 = Nothing

            Dim edmSearch As IEdmSearch7 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_Search), IEdmSearch7)
            edmSearch.FindFiles = True
            edmSearch.FindFolders = False
            edmSearch.Recursive = True
            edmSearch.SetToken(EdmSearchToken.Edmstok_FindItems, True)
            edmSearch.StartFolderID = edmVault.ItemRootFolderID
            edmSearch.FileName = PartNumber & ".<item>"

            Dim edmResult As IEdmSearchResult5 = edmSearch.GetFirstResult()
            If edmResult Is Nothing Then Throw New Exception("Item Search Returned Nothing")

            While Not edmResult Is Nothing
                If edmResult.IsKindOf(EdmObjectType.EdmObject_Item) Then
                    edmFile = DirectCast(edmResult, IEdmFile8)
                    Exit While
                End If
                edmResult = edmSearch.GetNextResult()
            End While

            edmResult = Nothing
            If edmFile Is Nothing Then Return -1
            Return edmFile.ID

        Catch ex As Exception
            Throw New Exception("Failed To Find Item in EPDM", ex)
        End Try
    End Function

    ''' <summary>
    ''' Searches The EPDM Vault for a part and returns the attached cad document throws an exception if no cad documents are attached
    ''' </summary>
    ''' <param name="PartNumber">Part number To Find</param>
    ''' <returns>Path to parts referenced document</returns>
    ''' <remarks></remarks>
    Public Function FindManagedItem(ByVal PartNumber As String) As String
        Try

            Dim edmFile As IEdmFile8 = Nothing

            Dim edmSearch As IEdmSearch7 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_Search), IEdmSearch7)
            edmSearch.FindFiles = True
            edmSearch.FindFolders = False
            edmSearch.Recursive = True
            edmSearch.SetToken(EdmSearchToken.Edmstok_FindItems, True)
            edmSearch.StartFolderID = edmVault.ItemRootFolderID
            edmSearch.FileName = PartNumber & ".<item>"

            Dim edmResult As IEdmSearchResult5 = edmSearch.GetFirstResult()
            If edmResult Is Nothing Then Throw New Exception("Item Search Returned Nothing")

            While Not edmResult Is Nothing
                If edmResult.IsKindOf(EdmObjectType.EdmObject_Item) Then
                    edmFile = DirectCast(edmResult, IEdmFile8)
                    Exit While
                End If
                edmResult = edmSearch.GetNextResult()
            End While

            edmResult = Nothing
            If edmFile Is Nothing Then Throw New Exception("Failed To Find A Classified Item")

            Dim edmRefTree As IEdmReference8 = DirectCast(edmFile.GetReferenceTree(edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID), IEdmReference8)
            If edmRefTree Is Nothing Then Throw New Exception("Failed To Get Reference Tree For The Item")

            Dim ProjName As String = ""

            'get any cad data attached to the item
            Dim RefdCADDoc As String = Nothing

            Dim edmPos As IEdmPos5 = edmRefTree.GetFirstChildPosition2(ProjName, True, False, EdmRefFlags.EdmRef_File + EdmRefFlags.EdmRef_Dynamic + EdmRefFlags.EdmRef_Static)
            While Not edmPos.IsNull
                Dim edmFileRef As IEdmReference8 = DirectCast(edmRefTree.GetNextChild(edmPos), IEdmReference8)
                If edmFileRef.ReferencedAs.ToUpper().EndsWith(".SLDASM") Or edmFileRef.ReferencedAs.ToUpper().EndsWith(".SLDPRT") Then
                    RefdCADDoc = edmFileRef.ReferencedAs.ToUpper().Replace("<ITEMS>", "")
                    RefdCADDoc = RefdCADDoc.Replace("\\", "\")
                End If
                edmFileRef = Nothing
            End While
            edmRefTree = Nothing
            edmFile = Nothing
            If RefdCADDoc Is Nothing Then Throw New Exception("Failed To Fetch Referenced Document")
            Return RefdCADDoc
        Catch ex As Exception
            Throw New Exception("Failed To Find Item in EPDM", ex)
        End Try
    End Function

    ''' <summary>
    ''' Renames A File in the EPDM Vault given the id and the new name
    ''' </summary>
    ''' <param name="FileID">ID of the file to rename</param>
    ''' <param name="NewName">New name for the file, just the name without path or extension</param>
    ''' <remarks></remarks>
    Public Sub RenameFile(ByVal FileID As Integer, ByVal NewName As String)
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            Dim FileExtension As String = IO.Path.GetExtension(edmFile.Name)
            If Not FileExtension.StartsWith(".") Then FileExtension = "." & FileExtension
            edmFile.RenameEx(0, NewName & FileExtension, 0)
        Catch ex As Exception
            Throw New Exception("Failed Rename EPDM File", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Sets The Current Revision Of a File
    ''' </summary>
    ''' <param name="FileID">File To Change the revision on</param>
    ''' <param name="RevisionCounterName">Name of the Revision Counter To Update</param>
    ''' <param name="RevisionCount">New Value To set the revision counter to</param>
    ''' <remarks></remarks>
    Public Sub SetFileRevision(ByVal FileID As Integer, ByVal RevisionCounterName As String, ByVal RevisionCount As Integer)
        Try
            If RevisionCount <= 0 Then Return
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            edmFile.UnlockFile(0, "", EdmUnlockFlag.EdmUnlock_Simple)
            'edmFile.IncrementRevision(0, edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID, "")
            Dim edmRevMgr As IEdmRevisionMgr3 = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_RevisionMgr), IEdmRevisionMgr3)
            Dim RevisionCounters As List(Of EdmRevCounter) = New List(Of EdmRevCounter)
            Dim RevCounter As EdmRevCounter
            RevCounter.mbsComponentName = "Controlled Design Numeric" 'RevisionCounterName
            RevCounter.mlCounter = RevisionCount
            RevisionCounters.Add(RevCounter)
            edmRevMgr.SetRevisionCounters(FileID, RevisionCounters.ToArray())

            edmRevMgr.IncrementRevision(FileID)
            Dim Errors As Array = Nothing
            edmRevMgr.Commit("Auto Revision Import From WM", Errors)
            If Errors.Length > 0 Then
                Throw New Exception("Error Setting Revision in File")
            End If
            edmRevMgr = Nothing
            edmFile.LockFile(edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID, 0)

            edmFile = Nothing
        Catch ex As Exception
            Throw New Exception("Failed Setting Revision of File", ex)
        End Try
    End Sub


    ''' <summary>
    ''' Checks In the file to the EPDM Vault
    ''' </summary>
    ''' <param name="FileID">EPDM ID of The File To Check In</param>
    ''' <remarks>Modified by Tim Cutler and Simon Turner</remarks>
    Public Sub CheckInFile(ByVal FileID As Integer)
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            If Not edmFile.IsLocked Then Return
            edmFile.UnlockFile(0, "Checked in by Ontology Studio", EdmUnlockFlag.EdmUnlock_Simple)
            'this might be quicker
            'Dim edmBatchUnlock As IEdmBatchUnlock = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock), IEdmBatchUnlock)
            'Dim selectedItems(0) As EdmSelItem
            'selectedItems(0).mlDocID = FileID
            'selectedItems(0).mlProjID = edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID
            'edmBatchUnlock.AddSelection(edmVault, DirectCast(selectedItems, Array))
            'edmBatchUnlock.CreateTree(0, EdmUnlockBuildTreeFlags.Eubtf_MayUnlock)
            'edmBatchUnlock.UnlockFiles(0)
        Catch ex As Exception
            Throw New Exception("Failed Unlocking File", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Checks out a file to the EPDM Vault
    ''' </summary>
    ''' <param name="FileID"></param>
    ''' <remarks>Addition by Tim</remarks>
    Public Sub CheckOutFile(ByVal FileID As Integer)
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            Dim edmFileFolderID As Integer = edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID
            If edmFile.IsLocked Then Return
            edmFile.LockFile(edmFileFolderID, 0, EdmLockFlag.EdmLock_Simple)
        Catch ex As Exception
            Throw New Exception("Failed Locking File", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Move file to next state
    ''' </summary>
    ''' <param name="FileID"></param>
    ''' <param name="StateName"></param>
    ''' <remarks></remarks>
    Public Sub MoveFileToState(ByVal FileID As Integer, ByVal StateName As String)
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            Dim edmFileFolderID As Integer = edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID
            edmFile.ChangeState(DirectCast(StateName, Object), edmFileFolderID, "Initialization", 0, EdmStateFlags.EdmState_Simple)
        Catch ex As Exception
            Throw New Exception("Failed changing state", ex)
        End Try
    End Sub


    ''' <summary>
    ''' Adds A Reference Between an Item in EPDM and A File In EPDM
    ''' </summary>
    ''' <param name="FileId">ID of the File To Reference</param>
    ''' <param name="ItemId">ID of the Item To Update</param>
    ''' <param name="Configuration">Configuration to Reference Empty string if no cofiguration</param>
    ''' <returns>Error or success string</returns>
    ''' <remarks>Modified by Tim Cutler and Simon Turner</remarks>
    Public Function ReferenceFileToItem(ByVal FileId As Integer, ByVal ItemId As Integer, ByVal Configuration As String) As String
        Dim File As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileId), IEdmFile8)
        File.GetFileCopy(0)
        Dim AddRefs(0) As EdmItemRef
        Dim OldRefs() As EdmItemRef = Nothing
        AddRefs(0).moParentNamePathOrItemID = ItemId
        AddRefs(0).moNamePathOrID = FileId
        AddRefs(0).mbsConfiguration = Configuration
        AddRefs(0).mlEdmRefFlags = EdmRefFlags.EdmRef_File
        Dim edmBatchAddRef As IEdmBatchItemReferenceUpdate = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchItemReferenceUpdate), IEdmBatchItemReferenceUpdate)
        edmBatchAddRef.UpdateReferences(DirectCast(AddRefs, System.Array), DirectCast(OldRefs, System.Array))
        edmBatchAddRef = Nothing
        Dim ErrStr As String = edmVault.GetErrorName(AddRefs(0).mhResult)
        Return ErrStr
    End Function


    ''' <summary>
    ''' Returns The BOM for A Solidowrks assembly in the EPDM Vault
    ''' </summary>
    ''' <param name="FileID">ID of The File in the EPDM Vault</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAssemblyBOM(ByVal FileID As Integer, ByVal Configuration As String, ByVal BomLayoutName As String) As BillOfMaterials
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            If edmFile Is Nothing Then Throw New Exception("Failed To Get File Reference For this Item")

            Dim edmBomManager As IEdmBomMgr = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BomMgr), IEdmBomMgr)
            Dim edmBomLayouts As Array = Nothing
            edmBomManager.GetBomLayouts(edmBomLayouts)
            Dim LayoutID As Integer = -1
            For Each BomLayout As EdmBomLayout In edmBomLayouts
                If String.Compare(BomLayout.mbsLayoutName, BomLayoutName, True) = 0 Then
                    LayoutID = BomLayout.mlLayoutID
                    Exit For
                End If
            Next
            If LayoutID < 0 Then
                Throw New Exception("Failed To Find Solidworks BOM Layout")
            End If
            Dim BomView As IEdmBomView2 = DirectCast(edmFile.GetComputedBOM(LayoutID, -1, Configuration, EdmBomFlag.EdmBf_AsBuilt), IEdmBomView2)
            Dim BomColumns As Array = Nothing
            BomView.GetColumns(BomColumns)

            Dim BomRows As Array = Nothing
            BomView.GetRows(BomRows)
            Dim PartNumber As String = ""
            Dim Count As Integer = 0
            Dim SWBOM As BillOfMaterials = New BillOfMaterials()

            For Each BomCell As EdmBomCell In BomRows
                For Each Column As EdmBomColumn In BomColumns
                    If String.Compare(Column.mbsCaption, "PART NUMBER", True) = 0 Then
                        Dim Value As Object = Nothing
                        Dim ComputedValue As Object = Nothing
                        Dim ReadOnlyFlag As Boolean = False
                        Dim BomConfiguration As String = Configuration
                        BomCell.GetVar(Column.mlVariableID, Column.meType, Value, ComputedValue, BomConfiguration, ReadOnlyFlag)
                        PartNumber = DirectCast(Value, String)
                    ElseIf String.Compare(Column.mbsCaption, "QTY", True) = 0 Then
                        Dim Value As Object = Nothing
                        Dim ComputedValue As Object = Nothing
                        Dim ReadOnlyFlag As Boolean = False
                        Dim BomConfiguration As String = Configuration
                        BomCell.GetVar(Column.mlVariableID, Column.meType, Value, ComputedValue, BomConfiguration, ReadOnlyFlag)
                        Count = Integer.Parse(DirectCast(Value, String))
                    End If
                Next
                If Count > 0 Then
                    SWBOM.Add(New BOMItem(PartNumber, Count))
                End If
            Next

            BomView = Nothing
            edmBomManager = Nothing
            edmFile = Nothing
            Return SWBOM

        Catch ex As Exception
            Throw New Exception("Failed Fetching Solidworks BOM", ex)
        End Try
    End Function

    ''' <summary>
    ''' Creates An Item BOM on an Item
    ''' </summary>
    ''' <param name="ParentID">ID of the item to attach the Bom</param>
    ''' <param name="ItemIDs">IDs of items to add to the bom</param>
    ''' <param name="ItemQuantity">A cross refernce of part number and quantities</param>
    ''' <param name="BomLayoutName">The name of the Bom layout</param>
    ''' <remarks></remarks>
    Public Sub CreateItemBom(ByVal ParentID As Integer, ByVal ItemIDs As List(Of Integer), ByVal ItemQuantity As Hashtable, ByVal BomLayoutName As String)
        Try
            If ItemIDs.Count <= 0 Then Return 'there is no bom
            Dim ParentItem As IEdmItem = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_Item, ParentID), IEdmItem)

            Dim AddItems As List(Of EdmItemRef) = New List(Of EdmItemRef)
            For Each IDToAdd As Integer In ItemIDs
                Dim AddItem As EdmItemRef = New EdmItemRef()
                AddItem.mbsConfiguration = ""
                AddItem.mlEdmRefFlags = EdmRefFlags.EdmRef_Item
                AddItem.moNamePathOrID = IDToAdd
                AddItem.moParentNamePathOrItemID = ParentID
                AddItems.Add(AddItem)
            Next
            Dim AddItemsArray As Array = DirectCast(AddItems.ToArray(), Array)
            Dim RemoveItemsArray As Array = Nothing

            ParentItem.UpdateReferences(AddItemsArray, RemoveItemsArray)
            If AddItemsArray IsNot Nothing Then
                For Each Added As EdmItemRef In AddItemsArray
                    If Added.mhResult <> 0 Then
                        Throw New Exception("An Unexpected Error Occured adding new items to the BOM")
                    End If
                Next
            End If

            'have now added all the necessary references
            'now need to update all the quantities
            'need to check in the item and check out again to update the references
            Dim UnlockUtility As IEdmBatchUnlock = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock), IEdmBatchUnlock)
            Dim GetUtility As IEdmBatchGet = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchGet), IEdmBatchGet)
            Dim Items(0) As EdmSelItem

            Items(0) = New EdmSelItem()
            Items(0).mlDocID = ParentItem.ID
            Dim ParentFile As IEdmFile8 = DirectCast(ParentItem, IEdmFile8)
            Items(0).mlProjID = ParentFile.GetNextFolder(ParentFile.GetFirstFolderPosition()).ID

            UnlockUtility.AddSelection(DirectCast(edmVault, EdmVault5), DirectCast(Items, Array))
            UnlockUtility.CreateTree(0, EdmUnlockBuildTreeFlags.Eubtf_MayUnlock)
            UnlockUtility.UnlockFiles(0)
            GetUtility.AddSelection(DirectCast(edmVault, EdmVault5), DirectCast(Items, Array))
            GetUtility.CreateTree(0, EdmGetCmdFlags.Egcf_Lock)
            GetUtility.GetFiles(0, Nothing)

            Dim edmBomManager As IEdmBomMgr = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BomMgr), IEdmBomMgr)
            Dim edmBomLayouts As Array = Nothing
            edmBomManager.GetBomLayouts(edmBomLayouts)
            Dim LayoutID As Integer = -1
            For Each BomLayout As EdmBomLayout In edmBomLayouts
                If String.Compare(BomLayout.mbsLayoutName, BomLayoutName, True) = 0 Then
                    LayoutID = BomLayout.mlLayoutID
                    Exit For
                End If
            Next
            If LayoutID < 0 Then
                Throw New Exception("Failed To Find Solidworks BOM Layout")
            End If
            Dim BomView As IEdmBomView2 = DirectCast(ParentFile.GetComputedBOM(LayoutID, -1, "", EdmBomFlag.EdmBf_AsBuilt), IEdmBomView2)
            Dim BomRows As Array = Nothing
            BomView.GetRows(BomRows)

            Dim BomColumns As Array = Nothing
            BomView.GetColumns(BomColumns)

            Dim RefCountColID As Integer = -1
            Dim NameCol As EdmBomColumn = Nothing

            For Each Column As EdmBomColumn In BomColumns
                If Column.meType = EdmBomColumnType.EdmBomCol_RefCount Then
                    RefCountColID = Column.mlVariableID
                ElseIf String.Compare(Column.mbsCaption, "PART NUMBER", True) = 0 Then
                    NameCol = Column
                End If
            Next

            For Each Cell As IEdmBomCell In BomRows
                Dim Value As Object = Nothing
                Dim ComputedValue As Object = Nothing
                Dim ReadOnlyFlag As Boolean = False
                Cell.GetVar(NameCol.mlColumnID, NameCol.meType, Value, ComputedValue, "", ReadOnlyFlag)
                Dim PartNumber As String = Value.ToString()
                Dim UpdatedQuantity As Integer = DirectCast(ItemQuantity(PartNumber.ToUpper()), Integer)
                Dim errmsg As String = ""
                If Not Cell.SetVar(RefCountColID, EdmBomColumnType.EdmBomCol_RefCount, UpdatedQuantity, "", EdmBomSetVarOption.EdmBomSetVarOption_Both, errmsg) Then
                    Throw New Exception("Failed To Update Quantity: " & errmsg)
                End If
            Next
            Dim err As String = ""
            Dim node As Integer = 0
            If BomView.Commit("", err, node) <> 0 Then
                Throw New Exception("Failed To Commit Quantity Changes: " & err)
            End If

        Catch ex As Exception
            Throw New Exception("Failed Creating an Item Bom", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Gets The Current Local Version Number of a selected File
    ''' </summary>
    ''' <param name="FileID">ID of the File</param>
    ''' <returns>If a local version exists then the current version number -1 if no local version exists</returns>
    ''' <remarks></remarks>
    Public Function GetLocalVersion(ByVal FileID As Integer) As Integer
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            Dim LocalVersion As Integer = edmFile.GetLocalVersionNo(edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID)
            edmFile = Nothing
            Return LocalVersion
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Gets A local Copy of a particular Version of the given file
    ''' </summary>
    ''' <param name="FileID">ID of the File</param>
    ''' <param name="Version">A valid Version of the File To obtain, 0 will obtain the latest visible version</param>
    ''' <remarks></remarks>
    Public Sub FetchVersion(ByVal FileID As Integer, ByVal Version As Integer)
        Try
            Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
            Dim edmBatchGet As IEdmBatchGet = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchGet), IEdmBatchGet)
            edmBatchGet.AddSelectionEx(DirectCast(edmVault, EdmVault5), FileID, edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID, Version)
            edmBatchGet.CreateTree(0, EdmGetCmdFlags.Egcf_Nothing)
            edmBatchGet.GetFiles(0)
            Dim test As EdmSelectionList5 = edmBatchGet.GetFileList(EdmGetFileListFlag.Egflf_GetFailed)
            edmFile = Nothing
            edmBatchGet = Nothing
        Catch ex As Exception
            Throw New Exception("Failed Getting Local Version", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Checks Out A List Of files
    ''' </summary>
    ''' <param name="FileIDs">List Of Files To check Out</param>
    ''' <remarks></remarks>
    Public Sub CheckOutFiles(ByVal FileIDs As List(Of Integer))
        Try
            Dim edmBatchGet As IEdmBatchGet = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchGet), IEdmBatchGet)
            Dim selItemList As List(Of EdmSelItem) = New List(Of EdmSelItem)
            For Each FileID As Integer In FileIDs
                Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
                Dim selItem As EdmSelItem = New EdmSelItem()
                selItem.mlDocID = FileID
                selItem.mlProjID = edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID
                selItemList.Add(selItem)
                edmFile = Nothing
            Next
            Dim selItemArr As Array = DirectCast(selItemList.ToArray(), Array)
            edmBatchGet.AddSelection(DirectCast(edmVault, EdmVault5), selItemArr)
            edmBatchGet.CreateTree(0, EdmGetCmdFlags.Egcf_Lock)
            edmBatchGet.GetFiles(0)
            edmBatchGet = Nothing
        Catch ex As Exception
            Throw New Exception("Failed Checking Out Files", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Checks In A List Of files
    ''' </summary>
    ''' <param name="FileIDs">List Of Files To check In</param>
    ''' <param name="Undo">If true all changes will de disregarded (undo checkout)</param>
    ''' <remarks></remarks>
    Public Sub CheckInFiles(ByVal FileIDs As List(Of Integer), ByVal Undo As Boolean)
        Try
            Dim edmBatchUnlock As IEdmBatchUnlock = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock), IEdmBatchUnlock)
            Dim selItemList As List(Of EdmSelItem) = New List(Of EdmSelItem)
            For Each FileID As Integer In FileIDs
                Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
                Dim selItem As EdmSelItem = New EdmSelItem()
                selItem.mlDocID = FileID
                selItem.mlProjID = edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID
                selItemList.Add(selItem)
                edmFile = Nothing
            Next
            Dim selItemArr As Array = DirectCast(selItemList.ToArray(), Array)
            edmBatchUnlock.AddSelection(DirectCast(edmVault, EdmVault5), selItemArr)
            If Not Undo Then
                edmBatchUnlock.CreateTree(0, EdmUnlockBuildTreeFlags.Eubtf_MayUnlock)
            Else
                edmBatchUnlock.CreateTree(0, EdmUnlockBuildTreeFlags.Eubtf_MayUndoLock + EdmUnlockBuildTreeFlags.Eubtf_UndoLockDefault)
            End If
            edmBatchUnlock.UnlockFiles(0)
            edmBatchUnlock = Nothing
        Catch ex As Exception
            Throw New Exception("Failed Checking Out Files", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Chanegs the state of a list of files
    ''' </summary>
    ''' <param name="FileIds">List of files to transition</param>
    ''' <param name="TransitionID">The EPDM ID of the transition to follow</param>
    ''' <remarks></remarks>
    Public Sub ChangeStateFiles(ByVal FileIds As List(Of Integer), ByVal TransitionID As Integer)
        Try
            Dim edmBatchChangeState As IEdmBatchChangeState = DirectCast(edmVault.CreateUtility(EdmUtility.EdmUtil_BatchChangeState), IEdmBatchChangeState)
            For Each FileID As Integer In FileIds
                Dim edmFile As IEdmFile8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_File, FileID), IEdmFile8)
                edmBatchChangeState.AddFile(FileID, edmFile.GetNextFolder(edmFile.GetFirstFolderPosition()).ID)
                edmFile = Nothing
            Next
            Dim edmTransition As IEdmTransition8 = DirectCast(edmVault.GetObject(EdmObjectType.EdmObject_Transition, TransitionID), IEdmTransition8)
            If Not edmBatchChangeState.CreateTree(edmTransition.Name) Then
                edmTransition = Nothing
                edmBatchChangeState = Nothing
                Return
            End If
            edmBatchChangeState.ChangeState(0)
            edmBatchChangeState = Nothing
        Catch ex As Exception
            Throw New Exception("Failed Changing State Of Files", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Converts the specified string representation of a date from WorkManager to a datetime equivalent using the specified format
    ''' </summary>
    ''' <param name="WmDate">The date string</param>
    ''' <param name="DateFormat">A string format code : - 'UTC' to convert from a Workmanager UTC time integer 
    ''' and standard .Net date time formats http://msdn.microsoft.com/en-us/library/8kb3ddd4(v=vs.71).aspx </param>
    ''' <param name="Result"></param>
    ''' <returns>Indicates if the method was successful</returns>
    ''' <remarks></remarks>
    Private Function TryParseWmDate(ByVal WmDate As String, ByVal DateFormat As String, ByRef Result As DateTime) As Boolean
        Dim Success As Boolean = False
        If DateFormat.ToUpper = "UTC" Then
            Dim Seconds As Integer = 0
            If Integer.TryParse(WmDate, Seconds) Then
                Dim Epoch As DateTime = New DateTime(1970, 1, 1)
                Result = Epoch.AddSeconds(Seconds).ToLocalTime
                Success = True
            End If
        Else
            Success = DateTime.TryParseExact(WmDate, DateFormat, New Globalization.CultureInfo("en-GB"), Globalization.DateTimeStyles.AssumeLocal, Result)
        End If
        Return Success
    End Function

    Public ReadOnly Property VaultRootFolder() As String
        Get
            Return edmVault.RootFolderPath
        End Get
    End Property

    Public ReadOnly Property VaultItemRootFolder() As String
        Get
            Return edmVault.ItemRootFolder.LocalPath
        End Get
    End Property

    Public ReadOnly Property VaultItemRootFolderID() As Integer
        Get
            Return edmVault.ItemRootFolderID
        End Get
    End Property

    Public ReadOnly Property EpdmVault As IEdmVault12
        Get
            Return edmVault
        End Get
    End Property

End Class
