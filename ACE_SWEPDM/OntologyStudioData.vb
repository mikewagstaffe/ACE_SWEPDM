Public Class OntologyStudioData
    'Data coming back from ontology Studio
    Public Name As String = ""
    Public Description As String = ""
    Public RelativePath As String = ""

    'Data Going to ontology studio
    Public OriginalDescription As String = ""
    Public OriginalName As String = ""
    Public AssociatedFileName As String = ""
    Public ScreenShotImgPath As String = ""
    Public ConfigurationName As String = ""

    'Rename Flag
    Public RenameRequired As Boolean = False
    'Used to indicate that this part was already classified and the returned value from ontology studio is a
    'classified part already existing
    Public ReplacementPart As Boolean = False
End Class