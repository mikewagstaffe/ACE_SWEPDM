Imports System.Collections.Generic

Public Class BillOfMaterials
    Public BomList As List(Of BOMItem) = New List(Of BOMItem)
    Private SW_BomList As List(Of BOMItem) = New List(Of BOMItem)
    Private ItemBom As List(Of BOMItem) = New List(Of BOMItem)

    Public Sub Add(ByVal Part As BOMItem)
        BomList.Add(Part)
    End Sub

    Public Sub AddSW(ByVal SWBOM As BillOfMaterials)
        SW_BomList = SWBOM.BomList
    End Sub

    Public Sub CreateItemBom(ByRef ConnectionToOntologyStudio As OntologyStudioConnection)
        Try
            For Each BomLine As BOMItem In BomList
                Dim OntologyStudioPart As String = ConnectionToOntologyStudio.LookupInOntologyStudio(BomLine.PartNumber)
                If OntologyStudioPart IsNot Nothing Then
                    BomLine.PartNumber = OntologyStudioPart
                    BomLine.IsClassifed = True
                End If
                Dim SWBomCount As Integer = 0
                For Each SWBomLine As BOMItem In SW_BomList
                    If String.Compare(BomLine.PartNumber, SWBomLine.PartNumber, True) = 0 Then
                        'the part in the bom exists in the solidworks bom
                        'dont add  this to the item bom
                        SWBomCount = SWBomLine.Quantity
                        Exit For
                    End If
                Next
                If BomLine.Quantity - SWBomCount > 0 Then
                    ItemBom.Add(BomLine.Copy())
                End If
            Next
            'After this ItemBOM will contain a correct list of parts that should be added as an Item bom
        Catch ex As Exception
            Throw New Exception("Failed Merging Solidworks BOM", ex)
        End Try
    End Sub

#If WORKMGR Then
    Public Sub ImportNonClassifedParts(ByRef ConnectionMgr As ConnectionManager)
        Try
            For Index As Integer = 0 To ItemBom.Count - 1
                If ItemBom(Index).IsClassifed Then Continue For
                Dim ItemToImport As WorkManagerPart = New WorkManagerPart(ConnectionMgr, ItemBom(Index).PartNumber)
                ItemToImport.ImportPart()
                Dim PartNumber As String = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.OntologyStudio), OntologyStudioConnection).LookupInOntologyStudio(ItemBom(Index).PartNumber)
                If PartNumber Is Nothing Then Throw New Exception("Failed importing A Part Clasification check returned Nothing")
                ItemBom(Index).PartNumber = PartNumber
                ItemBom(Index).IsClassifed = True
            Next
        Catch ex As Exception
            Throw New Exception("Failed Importing BOM Parts", ex)
        End Try
    End Sub
#End If

    ''' <summary>
    ''' Creates an Item BOM
    ''' </summary>
    ''' <param name="ParentID">ID of the parent item for this BOM</param>
    ''' <param name="ConnectionMgr">The connection manager </param>
    ''' <remarks></remarks>
    Public Sub CreateEPDMBom(ByVal ParentID As Integer, ByRef ConnectionMgr As ConnectionManager)
        Try
            Dim BomIDs As List(Of Integer) = New List(Of Integer)
            Dim ItemQuantity As Hashtable = New Hashtable()
            For Index As Integer = 0 To ItemBom.Count - 1
                Dim ItemID As Integer = DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).FindItem(ItemBom(Index).PartNumber)
                BomIDs.Add(ItemID)
                ItemQuantity.Add(ItemBom(Index).PartNumber.ToUpper(), ItemBom(Index).Quantity)
            Next
            DirectCast(ConnectionMgr.GetConnection(ConnectionMgrConnection.EPDM), EPDMConnection).CreateItemBom(ParentID, BomIDs, ItemQuantity, ConnectionMgr.EPDM_ConnectionInfo.BomLayoutName)
        Catch ex As Exception
            Throw New Exception("Failed Creating Item BOM", ex)
        End Try
    End Sub
End Class
