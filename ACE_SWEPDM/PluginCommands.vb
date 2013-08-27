' (C) Copyright 2013 by M.Wagstaffe (Ishida Europe Ltd)
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(ACE_SWEPDM.PluginCommands))> 
Namespace ACE_SWEPDM

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class PluginCommands
        'The Current Verision of the application
        Private Const BuildVersion As String = "V0.2.0"

        ' The CommandMethod attribute can be applied to any public  member 
        ' function of any public class.
        ' The function should take no arguments and return nothing.
        ' If the method is an instance member then the enclosing class is 
        ' instantiated for each document. If the member is a static member then
        ' the enclosing class is NOT instantiated.
        '
        ' NOTE: CommandMethod has overloads where you can provide helpid and
        ' context menu.

        ' Modal Command with localized name
        ' AutoCAD will search for a resource string with Id "MyCommandLocal" in the 
        ' same namespace as this command class. 
        ' If a resource string is not found, then the string "MyLocalCommand" is used 
        ' as the localized command name.
        ' To view/edit the resx file defining the resource strings for this command, 
        ' * click the 'Show All Files' button in the Solution Explorer;
        ' * expand the tree node for myCommands.vb;
        ' * and double click on myCommands.resx

        <CommandMethod("EPDM_CREATEPROJECT", CType((CommandFlags.Modal + CommandFlags.Session), CommandFlags))> _
        Public Sub NewElectricalProject()
            If ACECommands.ConnectionMgr Is Nothing Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Enterprise PDM Connection is not Initialised")
            End If
            ACECommands.NewElectricalProject()
        End Sub


        <CommandMethod("EPDM_COPYPROJECT", CType((CommandFlags.Modal + CommandFlags.Session), CommandFlags))> _
        Public Sub CopyElectricalProject()
            If ACECommands.ConnectionMgr Is Nothing Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Enterprise PDM Connection is not Initialised")
            End If
            ACECommands.NewCopyOfElectricalProject()
        End Sub




        '##################################################################################################################
        '############################## All Code below here is default template code ######################################

        '<CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)> _
        'Public Sub MyCommand() ' This method can have any name
        '    ' Put your command code here
        'End Sub

        '' Modal Command with pickfirst selection
        '<CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CType(CommandFlags.Modal + CommandFlags.UsePickSet, CommandFlags))> _
        'Public Sub MyPickFirst() ' This method can have any name
        '    Dim result As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection()
        '    If (result.Status = PromptStatus.OK) Then
        '        ' There are selected entities
        '        ' Put your command using pickfirst set code here
        '    Else
        '        ' There are no selected entities
        '        ' Put your command code here
        '    End If
        'End Sub

        '' Application Session Command with localized name
        '<CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CType(CommandFlags.Modal + CommandFlags.Session, CommandFlags))> _
        'Public Sub MySessionCmd() ' This method can have any name
        '    ' Put your command code here
        'End Sub

        '' LispFunction is similar to CommandMethod but it creates a lisp 
        '' callable function. Many return types are supported not just string
        '' or integer.
        '<LispFunction("MyLispFunction", "MyLispFunctionLocal")> _
        'Public Function MyLispFunction(ByVal args As ResultBuffer) As ResultBuffer ' This method can have any name
        '    ' Put your command code here

        '    ' Return a value to the AutoCAD Lisp Interpreter
        '    Dim retVal As ResultBuffer = New ResultBuffer()
        '    Return retVal
        'End Function

    End Class

End Namespace