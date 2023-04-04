Imports System.IO
Imports DriveWorks
Imports DriveWorks.Components
Imports DriveWorks.EventFlow
Imports DriveWorks.Forms
Imports DriveWorks.Navigation
Imports DriveWorks.Reporting
Imports DriveWorks.Security
Imports DriveWorks.SolidWorks.Components
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

<Task("Add Label", "embedded://DriveWorksPRGExtender.bomb.bmp", "PRG")>
Public Class addLabel
    Inherits Task

    Dim mLabelName As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Label Name", New FlowPropertyInfo("Name of the label to create", "Input"))
    Dim mLabelText As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Label Text", New FlowPropertyInfo("Text for the label ", "Input"))
    Dim mFormName As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Form Name", New FlowPropertyInfo("Form to add the label to", "Input"))
    Dim mLabelLeft As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Left", New FlowPropertyInfo("Left location (in pixels)", "Layout"))
    Dim mLabelTop As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Top", New FlowPropertyInfo("Top location (in pixels)", "Layout"))
    Dim mLabelHeight As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Height", New FlowPropertyInfo("Height (in pixels)", "Layout"))
    Dim mLabelWidth As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Width", New FlowPropertyInfo("Width (in pixels)", "Layout"))

    Public Sub New()
    End Sub

    Protected Overrides Sub Execute(ByVal ctx As SpecificationContext)
        ' ToDo: Validate inputs
        Dim labelName As String = mLabelName.Value
        Dim labelText As String = mLabelText.Value
        Dim formName As String = mFormName.Value
        Dim left As Integer = 0
        If Integer.TryParse(mLabelLeft.Value, left) Then
        End If
        Dim top As Integer = mLabelTop.Value
        If Integer.TryParse(mLabelTop.Value, top) Then
        End If
        Dim height As Integer = mLabelHeight.Value
        If Integer.TryParse(mLabelHeight.Value, height) Then
        End If
        Dim width As Integer = mLabelWidth.Value
        If Integer.TryParse(mLabelWidth.Value, width) Then
        End If
        Dim project = ctx.Project

        ' Try to get the form
        Dim form As FormNavigationStep = Nothing
        If project.Navigation.TryGetStep(formName, form) Then
            Dim ctrlColl As ControlCollection = form.Form.Controls
            ctrlColl.Add(GetType(Forms.Label), labelName)
            Dim label As Label = Nothing
            If project.Navigation.TryGetControl(Of Label)(labelName, label) Then
                label.Height = height
                label.Top = top
                label.Left = left
                label.Width = width
                label.Text = label.Text
            End If
        End If

        Me.SetState(NodeExecutionState.Successful)
    End Sub
End Class

<Task("Get Release Results", "embedded://DriveWorksPRGExtender.bomb.bmp", "PRG Tools", True)>
Public Class PRGGetModelList
    Inherits Task
    Private mInputValue As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Input Value", New FlowPropertyInfo("Models to be released", "Input"))
    Private mDeferredDrawings As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Deferred Drawings", New FlowPropertyInfo("Pipe-delimited list of drawings to be deferred", "Input"))
    Private mOverwrite As FlowProperty(Of Boolean) = Me.Properties.RegisterBooleanProperty("Overwrite Models", New FlowPropertyInfo("Overwrite existing models and drawings", "Settings"))

    Public Sub New()
        Me.Outputs.Register("Output", "Output value", GetType(String), "Output")
    End Sub

    Protected Overrides Sub Execute(ctx As SpecificationContext)
        Dim inputString As String = mInputValue.Value

        ' Create a release tracker object
        Dim relTracker = New ReleaseComponentReportTracker(ctx.Report)

        ' Create an environment
        Dim env As New ReleaseEnvironment() With {
            .Flags = If(mOverwrite.Value, ReleasedComponentFlags.ForceOverwrite, ReleasedComponentFlags.None),
            .OverwriteReleasedComponents = ctx.Environment.CanEditCompletedSpecifications AndAlso ctx.Environment.OverwriteReleasedComponents
        }

        ' Make sure that the deferred drawing list has no blanks and has an extension on all of the deferred drawing names, unless it's an asterisk
        Dim defDrwList = mDeferredDrawings.Value.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries).
                                                                Select(Function(drwName) As String
                                                                           If String.Equals(drwName, "*", StringComparison.Ordinal) OrElse
                                                                               Not String.IsNullOrEmpty(Path.GetExtension(drwName)) Then
                                                                               Return drwName
                                                                           End If
                                                                           Return String.Format("{0}.slddrw", drwName)
                                                                       End Function)
        Dim deferredDrawingList As String = String.Join("|", defDrwList)
        ' Perform the release and get the results
        Dim relResults = ReleaseComponentHelper.Release(env, ctx, inputString, relTracker, deferredDrawingList)

        ' Log the results in the database
        ctx.Group.ReleasedComponents.SaveReleaseResults(relResults)

        ' Parse the results to get the list of files
        Dim fileList As New List(Of String)
        For Each comp In relResults.Components
            ' Do we need to go through the LoopVariations? Do we need to recurse?
            If comp.LoopVariations.Count > 1 Then
                For Each loopComp In comp.LoopVariations
                    If loopComp.LoopVariations.Count > 1 Then
                        ' This is getting loopy
                    End If
                Next loopComp
            End If
            Dim swComp As ReleasedSolidWorksComponent = TryCast(comp, ReleasedSolidWorksComponent)
            If swComp IsNot Nothing Then fileList.Add(swComp.TargetPath)
        Next comp


        Dim outputString As String = String.Empty
        outputString = String.Join("|", fileList)
        FulfillOutput("Output", outputString)

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