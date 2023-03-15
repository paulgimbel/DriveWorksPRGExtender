Imports DriveWorks
Imports DriveWorks.EventFlow
Imports DriveWorks.Forms
Imports DriveWorks.Reporting
Imports DriveWorks.Security
Imports DriveWorks.Specification
Imports Titan.Rules.Execution

<Task("Choose Your Own Path", "embedded://DriveWorksPRGExtender.bomb.bmp", "PRG Tools", True)>
Public Class PRGSpecTasks
    Inherits Task
    Private mInputValue As FlowProperty(Of Int32) = Me.Properties.RegisterInt32Property("Input Value", New FlowPropertyInfo("A Number from 1 to 7 indicating the path to exit the task", "Input", {}))
    Private mOutputValue As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Output Value", New FlowPropertyInfo("Value to pass when exiting the task", "Output", {}))

    Public Sub New()
        Me.Outputs.Register("Path1", "First output path", GetType(String), "Path 1")
        Me.Outputs.Register("Path2", "First output path", GetType(String), "Path 2")
        Me.Outputs.Register("Path3", "First output path", GetType(String), "Path 3")
        Me.Outputs.Register("Path4", "First output path", GetType(String), "Path 4")
        Me.Outputs.Register("Path5", "First output path", GetType(String), "Path 5")
        Me.Outputs.Register("Path6", "First output path", GetType(String), "Path 6")
        Me.Outputs.Register("Path7", "First output path", GetType(String), "Path 7")
    End Sub

    Protected Overrides Sub Execute(ctx As SpecificationContext)
        Dim outputString As String = mOutputValue.Value

        If Not String.IsNullOrEmpty(outputString) And mInputValue.Value <= 7 And mInputValue.Value > 0 Then
            If FulfillOutput("Path" & mInputValue.Value, outputString) Then
                Me.SetState(NodeExecutionState.Successful, outputString)
            Else
                Me.SetState(NodeExecutionState.Failed, String.Format("No go. input:{0}, output{1}", outputString, mInputValue.Value))
            End If
        End If

    End Sub

    ''' <summary>
    ''' Internal function to populate an output node
    ''' </summary>
    ''' <param name="name">Name of the output node</param>
    ''' <param name="value">Value to pass out of the task</param>
    ''' <returns>True if successful</returns>
    Private Function FulfillOutput(name As String, value As Object) As Boolean
        Dim result As Boolean = False
        Try
            For index As Int16 = 0 To Me.Outputs.Count - 1
                If Me.Outputs.Item(index).Name = name Then
                    Me.Outputs.Item(index).Fulfill(value)
                    result = True
                    Exit For
                End If
            Next index
            Return result
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class

<Condition("Specification State", "embedded://DriveWorksPRGExtender.bomb.bmp")>
Public Class isInState
    Inherits Condition

    Private stateTitle As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("State Name", New FlowPropertyInfo("Name of the state to verify the current spec", "Input", {StandardRuleTypes.SpecificationFlowStateName}, True))

    Protected Overrides Function Evaluate(specificationContext As SpecificationContext) As Boolean
        If specificationContext.CurrentState.Title = stateTitle.Value Then
            Return True
        Else
            specificationContext.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Information, "Specification Task Condition", "Specification In State", "Specification was not in the specified state", String.Format("Specified State:{0}, Current State:{1}", stateTitle.Value, specificationContext.CurrentState.Title))
            Return False
        End If
    End Function
End Class

<Task("Preset preferences", "embedded://DriveWorksPRGExtender.bomb.bmp")>
Public Class setPrefs
    Inherits Task

    Dim mFileName As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("File Name", New FlowPropertyInfo("Name of the file to load", "Input"))

    Public Sub New()
        Me.Outputs.Register("FileOutput", "Contents of the file", GetType(String), "File Contents")
    End Sub

    Protected Overrides Sub Execute(ByVal ctx As SpecificationContext)
        Dim project = ctx.Project

        ' Get DriveWorks to figure out the full path to the file by just giving the file name
        ' and telling DriveWorks it's next to the project
        Dim fileName = project.Utility.ResolvePath(mFileName.Value, RelativeToDirectory.Project)
        Dim fileLines = System.IO.File.ReadAllLines(fileName)
        For index As Int16 = 0 To Me.Outputs.Count - 1
            If Me.Outputs.Item(index).Name = "FileOutput" Then
                Me.Outputs.Item(index).Fulfill(String.Join("|", fileLines))
                Exit For
            End If
        Next index

        For Each fileLine In fileLines

            ' Quick way to separate a line separated by an equals sign,
            ' NOTE: This will fail if there isn't an equals sign!
            Dim parts = fileLine.Split(New Char() {"="}, 2)
            Dim name = parts(0)
            Dim value = parts(1)

            Dim control As ControlBase = Nothing
            If project.Navigation.TryGetControl(name, control) Then
                control.SetInputValue(value)
            Else
                ctx.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Spec Task", "Preset preferences", "Control not found", String.Format("Unable to find control {0}", name))
                Me.SetState(NodeExecutionState.SuccessfulWithWarnings)
            End If
        Next
        Me.SetState(NodeExecutionState.Successful)
    End Sub
End Class


