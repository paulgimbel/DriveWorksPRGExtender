Imports System.Drawing.Drawing2D
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports DriveWorks
Imports DriveWorks.Components
Imports DriveWorks.Components.Tasks
Imports DriveWorks.Reporting
Imports DriveWorks.SolidWorks
Imports DriveWorks.SolidWorks.Generation
Imports DriveWorks.SolidWorks.Generation.Proxies
Imports DriveWorks.Specification.StandardTasks
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports Titan.Rules.Execution

Namespace DriveWorksPRGExtender

    <GenerationTask("Convert Part Model File",
                    "Import any non-native file and export it with SOLIDWORKS",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Parts, ComponentTaskSequenceLocation.All)>
    Public Class ConvertToDrive3D
        Inherits GenerationTask

        Private Const INPUT_FILENAME As String = "GenerationTask.inputName"
        Private Const OUTPUT_FILENAME As String = "GenerationTask.outputName"
        Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
            Get
                Return New ComponentTaskParameterInfo() {
                        New ComponentTaskParameterInfo(INPUT_FILENAME, "Input file Name", "Full path to the file to be imported"),
                        New ComponentTaskParameterInfo(OUTPUT_FILENAME, "Output file name", "Full path to the file to be exported (extension will determine format)")
                        }
            End Get
        End Property

        Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
            Dim swApp As SldWorks = model.Application.Instance
            Dim inputFilename As String = String.Empty
            Dim ouputFilename As String = String.Empty

            If Not Me.Data.TryGetParameterValue(INPUT_FILENAME, inputFilename) Then
                ' No entity was received to offset, can't go on
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "No input filename", "No value was received for the input file name")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(OUTPUT_FILENAME, ouputFilename) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "No output filename", "No value provided for the output file name")
                Return
            End If

            'Dim stepInfo As ImportStepData = swApp.GetImportFileData(inputFilename)
            Dim errors, warnings As Long
            Dim importedModel As ModelDoc2 = swApp.LoadFile4(inputFilename, "r", swApp.GetImportFileData(inputFilename), errors)
            If errors + warnings > 0 Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "Errors opening file", String.Format("Errors: {0}; Warnings: {1}", errors, warnings))
            Else
                ' Report success
            End If
            ' swApp.ActivateDoc2(System.IO.Path.GetFileNameWithoutExtension(importedModel.GetPathName) & ".SLDPRT", True, errors)
            If errors + warnings > 0 Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "Error activating file", String.Format("Errors: {0}; Warnings: {1}", errors, warnings))
            Else
                ' Report success
            End If
            Dim results As Integer = importedModel.SaveAs3(ouputFilename, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent)
            If errors + warnings > 0 Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "Errors saving file", String.Format("Errors: {0}; Warnings: {1}", errors, warnings))
            Else
                ' Report success
            End If

            swApp.CloseDoc(System.IO.Path.GetFileNameWithoutExtension(ouputFilename) & ".SLDPRT")
        End Sub
    End Class

    <GenerationTask("Drive View Style",
                "Set the shading style for a drawing view",
                "embedded://DriveWorksPRGExtender.ViewStyles.png",
                "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
    Public Class DriveViewStyle
        Inherits GenerationTask

        Private Const PARAM_NAME_VIEW_NAMES As String = "ViewNames"
        Private Const PARAM_NAME_VIEW_STYLE As String = "ViewStyles"

        Private Enum ViewStyle
            Wireframe = 0
            HiddenLinesVisible = 1
            HiddenLinesRemoved = 2
            Shaded = 3
            ShadedWithEdges = 4
        End Enum

        Private Const TASK_NAME As String = "Drive View Style"
        Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
            Get
                Return New ComponentTaskParameterInfo() {
                        New ComponentTaskParameterInfo(PARAM_NAME_VIEW_NAMES,
                                                       "Drawing View Names",
                                                       "Drawing view names to set",
                                                       "Inputs"),
                        New ComponentTaskParameterInfo(PARAM_NAME_VIEW_STYLE,
                                                       "Drawing View Style",
                                                       "Display style to set",
                                                       "Inputs",
                                                       New ComponentTaskParameterMetaData(Forms.DataModel.PropertyBehavior.StandardOptionDynamicManual, GetType(ViewStyle), ViewStyle.Shaded))}
            End Get
        End Property

        Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
            ' Validate Inputs
            Dim viewNamesRaw As String = String.Empty
            If Not Me.Data.TryGetParameterValue(PARAM_NAME_VIEW_NAMES, True, viewNamesRaw) Then
                Me.SetExecutionResult(TaskExecutionResult.Failed, "The property 'Drawing View Names' is invalid.")
                Return
            End If
            viewNamesRaw = viewNamesRaw.Trim()
            If String.IsNullOrWhiteSpace(viewNamesRaw) Then
                Me.SetExecutionResult(TaskExecutionResult.Failed, "The property 'Drawing View Names' cannot be blank.")
                Return
            End If

            Dim viewStyle As ViewStyle

            If Not Me.Data.TryGetParameterValueAsEnum(PARAM_NAME_VIEW_STYLE, viewStyle) Then
                Dim viewStateRaw = Me.Data.GetParameterValue(PARAM_NAME_VIEW_STYLE, String.Empty)
                If String.IsNullOrWhiteSpace(viewStateRaw) Then
                    Me.SetExecutionResult(TaskExecutionResult.Failed, "The property 'Drawing View Style' cannot be blank.")
                Else
                    Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("The value {0} is not valid for the property 'Drawing View Style'.", viewStateRaw))
                End If
                Return
            End If

            ' Convert Inputs To Lists
            Dim viewNamesToSet As List(Of String) = viewNamesRaw.Split("|").ToList
            Dim draw As DrawingDoc = Nothing
            Try

                draw = TryCast(model.Model, DrawingDoc)
                For Each viewName In viewNamesToSet
                    If draw IsNot Nothing Then
                        ' Select the drawing view by name
                        If model.Model.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) Then
                            Select Case viewStyle
                                Case ViewStyle.Shaded
                                    draw.ViewDisplayShaded()
                                Case ViewStyle.HiddenLinesRemoved
                                    draw.ViewDisplayHidden()
                                Case ViewStyle.HiddenLinesVisible
                                    draw.ViewDisplayHiddengreyed()
                                Case ViewStyle.Wireframe
                                    draw.ViewDisplayWireframe()
                                Case ViewStyle.ShadedWithEdges
                                    draw.ViewDisplayShaded()
                                    draw.ViewModelEdges()
                            End Select
                        Else
                            Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Task", TASK_NAME, "Unable to retrieve the drawing document", String.Format("Drawing view: {0}", viewName))
                        End If
                    Else
                        Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Task", TASK_NAME, "Unable to retrieve the drawing document", "")
                    End If
                    model.Model.ClearSelection2(True)
                Next
            Catch ex As Exception

                Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception: {0}", ex.Message))
            Finally 'Clean Up
                If draw IsNot Nothing Then
                    'Marshal.ReleaseComObject(draw)
                    draw = Nothing
                End If

            End Try
        End Sub
    End Class


    <GenerationTask("Add Revision",
                   "Add a Revision to the table and a revision symbol to the drawing",
                   "embedded://DriveWorksPRGExtender.Triangle.bmp",
                   "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.After)>
    Public Class AddRevision
        Inherits GenerationTask

        Private Const REV_LETTER As String = "revLetter"
        Private Const REV_DESC As String = "revDesc"
        Private Const REV_DESC_COL As String = "revDescriptionColumn"
        Private Const REV_LOC_X As String = "revLocX"
        Private Const REV_LOC_Y As String = "revLocY"

        Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
            Get
                Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(REV_LETTER, "Revision", "Letter or indicator to use in the revision table"),
                    New ComponentTaskParameterInfo(REV_DESC, "Revision Description", "Text to include in the revision table"),
                    New ComponentTaskParameterInfo(REV_DESC_COL, "Revision Description Column", "Column index for the description column"),
                    New ComponentTaskParameterInfo(REV_LOC_X, "X", "X location to place the revision symbol (in mm)"),
                    New ComponentTaskParameterInfo(REV_LOC_Y, "Y", "Y location to place the revision symbol (in mm)")
                    }
            End Get
        End Property

        Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
            ' Validate Inputs
            Dim revLetter As String = String.Empty
            Dim revDesc As String = String.Empty
            Dim revDescColString As String = String.Empty
            Dim revDescCol As Integer
            Dim revXVal As String = String.Empty
            Dim revYVal As String = String.Empty
            Dim revX, revY As Double
            Dim includeRevSymbol As Boolean = True

            If Not Me.Data.TryGetParameterValue(REV_LETTER, revLetter) Then
                ' No revision letter was received, can't go on
                Me.SetExecutionResult(TaskExecutionResult.Failed, "No value was received for the revision letter")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(REV_DESC, revDesc) Then
                Me.SetExecutionResult(TaskExecutionResult.Failed, "No value provided for the description")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(REV_DESC_COL, revDescColString) Or Not Integer.TryParse(revDescColString, revDescCol) Then
                Me.SetExecutionResult(TaskExecutionResult.Failed, "No valid value provided for the description column index")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(REV_LOC_X, revXVal) Or Not Double.TryParse(revXVal, revX) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "No X location", "No valid value provided for the X location of the revision symbol.")
                includeRevSymbol = False
            ElseIf Not Me.Data.TryGetParameterValue(REV_LOC_Y, revYVal) Or Not Double.TryParse(revYVal, revY) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", "Convert", "No Y location", "No valid value provided for the Y location of the revision symbol.")
                includeRevSymbol = False
            End If

            Dim swApp As SldWorks = Nothing
            Dim currentSheet As Sheet = Nothing
            Dim revTblFeat As RevisionTableFeature = Nothing
            Dim revTblAnn As RevisionTableAnnotation = Nothing
            Dim draw As DrawingDoc = Nothing
            Try
                swApp = model.Application.Instance
                draw = model.Drawing
                If draw Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.Failed, "Unable to retrieve the drawing")
                    Return
                End If
                currentSheet = draw.GetCurrentSheet()
                If currentSheet Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.Failed, "Unable to retrieve current sheet")
                    Return
                End If
                revTblAnn = currentSheet.RevisionTable
                If revTblAnn Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.Failed, "Unable to retrieve a revision table on the current sheet")
                    Return
                End If
                revTblFeat = revTblAnn.RevisionTableFeature
                If revTblFeat Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.Failed, "Unable to retrieve the revision table definition")
                    Return
                End If
                Dim newRevisionID As Long = revTblAnn.AddRevision(revLetter)
                If newRevisionID < 0 Then
                    Me.SetExecutionResult(TaskExecutionResult.Failed, "SOLIDWORKS failed to create the new revision")
                    Return
                Else
                    Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", "Add Revision", "New revision successfully added to table", String.Format("New ID: {0}, Revision: {1}", newRevisionID, revLetter))
                End If
                ' Add description to new revision row in table by recasting the RevisionTableAnnotation object as a generic TableAnnotation
                Dim tblAnn As TableAnnotation = TryCast(revTblAnn, TableAnnotation)
                If tblAnn Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Revision created, but unable to set description")
                Else
                    tblAnn.Text2(revTblAnn.GetRowNumberForId(newRevisionID), revDescCol - 1, True) = revDesc
                End If
                If includeRevSymbol Then
                    draw.InsertRevisionSymbol(revX / 1000, revY / 1000)
                End If
            Catch ex As Exception
                Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception: {0}", ex.Message))
            Finally
                If swApp IsNot Nothing Then
                    swApp = Nothing
                End If
                If currentSheet IsNot Nothing Then
                    currentSheet = Nothing
                End If
                If revTblFeat IsNot Nothing Then
                    revTblFeat = Nothing
                End If
                If revTblAnn IsNot Nothing Then
                    revTblAnn = Nothing
                End If
                If draw IsNot Nothing Then
                    draw = Nothing
                End If

                Me.SetExecutionResult(TaskExecutionResult.Success)
            End Try


        End Sub

    End Class

    <GenerationTask("Get List Of Models and Drawings",
                    "Obtain a list of all models and drawings to be generated from this generation context",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Assemblies, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After + ComponentTaskSequenceLocation.PreClose)>
    Public Class GetListOfModels
        Inherits GenerationTask

        Private Const OUTPUT_PATH As String = "GenerationTask.outputPath"

        Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
            Get
                Return New ComponentTaskParameterInfo() {
                        New ComponentTaskParameterInfo(OUTPUT_PATH, "Output Path", "Path to store the list (ex. \\MyServer\MyShare\Subfolder\filename.txt)")
                }
            End Get
        End Property

        Const taskName = "Get Model and Drawing List"

        Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)

        End Sub
    End Class

    <GenerationTask("Mate with Mate References",
                    "Select mate references from two components to mate them together",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Assemblies, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After + ComponentTaskSequenceLocation.PreClose)>
    Public Class MateWithMateReferences
        Inherits GenerationTask

        Private Const COMP1_NAME As String = "GenerationTask.component1"
        Private Const COMP1_REFS As String = "GenerationTask.comp1References"
        Private Const COMP2_NAME As String = "GenerationTask.component2"
        Private Const COMP2_REFS As String = "GenerationTask.comp2References"

        Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
            Get
                Return New ComponentTaskParameterInfo() {
                        New ComponentTaskParameterInfo(COMP1_NAME, "First Component", "Name of the first component, including instance (ex. Frame-1)"),
                        New ComponentTaskParameterInfo(COMP1_REFS, "First Component References", "Names of the mate references on the first component (ex. BaseMateRef|LeftHoleMateRef|CornerMateRef)"),
                        New ComponentTaskParameterInfo(COMP2_NAME, "Second Component", "Name of the second component, including instance (ex. Frame-1)"),
                        New ComponentTaskParameterInfo(COMP2_REFS, "Second Component References", "Names of the mate references on the second component (ex. BaseMateRef|LeftHoleMateRef|CornerMateRef)")
                        }
            End Get
        End Property

        Const taskName = "Insert Mate References"

        Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
            Dim swApp As SldWorks = model.Application.Instance
            Dim assy As AssemblyDoc = TryCast(model.Model, AssemblyDoc)
            Dim SelMgr As SelectionMgr = model.Model.SelectionManager
            If assy Is Nothing Then
                Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "Model is not an assembly", "SOLIDWORKS proxy model could not be recast as an assembly document")
                Return
            End If

            Dim comp1Name As String = String.Empty
            Dim comp1RefsString As String = String.Empty
            Dim comp2Name As String = String.Empty
            Dim comp2RefsString As String = String.Empty
            ' Verify that we got data in all of the inputs
            If Not Me.Data.TryGetParameterValue(COMP1_NAME, comp1Name) Then
                ' No entity was received to offset, can't go on
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "Component 1 not specified", "No value was received for the component 1 name")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(COMP1_REFS, comp1RefsString) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "No mate reference names provided for component 1", "No value provided for the mate references for component 1")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(COMP2_NAME, comp2Name) Then
                ' No entity was received to offset, can't go on
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "Component 2 not specified", "No value was received for the component 2 name")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(COMP2_REFS, comp2RefsString) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "No mate reference names provided for component 2", "No value provided for the mate references for component 2")
                Return
            End If

            'Get the components
            Dim swComp1 As Component2 = GetComponentByName(assy, comp1Name)
            If swComp1 Is Nothing Then
                Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "Unable to get component 1", String.Format("SOLIDWORKS component <{0}> could not be found", comp1Name))
                Return
            End If
            Dim swComp2 As Component2 = GetComponentByName(assy, comp2Name)
            If swComp2 Is Nothing Then
                Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "Unable to get component 2", String.Format("SOLIDWORKS component <{0}> could not be found", comp2Name))
                Return
            End If

            ' Split the Mate Reference inputs out in case they provided more than one
            Dim comp1Refs As List(Of String) = comp1RefsString.Split("|").ToList()
            Dim comp2Refs As List(Of String) = comp2RefsString.Split("|").ToList()
            If comp1Refs.Count <> comp2Refs.Count Then Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "Mate reference lists for components 1 and 2 are different lengths.", String.Format("Component 1 has <{0}> references while component 2 has <{1}>. The first <{2}> references in each list will be used.", comp1Refs.Count, comp2Refs.Count, Math.Min(comp1Refs.Count, comp2Refs.Count)))
            Dim AssyName As String = System.IO.Path.GetFileNameWithoutExtension(model.Model.GetPathName)

            ' Walk through the references in each Mate Ref
            For i As Int16 = 0 To Math.Min(comp1Refs.Count, comp2Refs.Count) - 1
                'Get the mate references
                Dim mr1 As MateReference = GetMateRefByName(swComp1, comp1Refs(i))
                Dim mr2 As MateReference = GetMateRefByName(swComp2, comp2Refs(i))

                'Walk through the reference entities/mates in each Mate Reference
                For j As Int16 = 0 To Math.Min(mr1.ReferenceEntityCount, mr2.ReferenceEntityCount) - 1
                    'Get MateType
                    Dim mr1Type As swMateReferenceType_e = mr1.ReferenceType(j)
                    Dim mr2Type As swMateReferenceType_e = mr2.ReferenceType(j)
                    ' Do not create the mate if the mate types don't match
                    If mr1Type = mr2Type Then

                        Dim selData As SelectData = SelMgr.CreateSelectData()

                        ' Select Reference 1
                        Dim mrRef1Type As swSelectType_e
                        mrRef1Type = mr1.ReferenceEntityType(j)
                        Dim mrRef1TypeString As String = GetEntityName(mrRef1Type)
                        Dim mrRef1 As SolidWorks.Interop.sldworks.IEntity = mr1.ReferenceEntity2(j)
                        If mrRef1 Is Nothing Then
                            Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Information, "Generation Task", taskName, "Reference for Component 1 could not be found", "")
                        Else
                            Dim corrEnt1 As Entity = swComp1.GetCorresponding(mrRef1)
                            If corrEnt1 IsNot Nothing Then
                                corrEnt1.Select2(False, 1)
                            Else
                            End If
                        End If

                        ' Select Reference 2
                        Dim mrRef2Type As swSelectType_e
                        mrRef2Type = mr2.ReferenceEntityType(j)
                        Dim mrRef2TypeString As String = GetEntityName(mrRef2Type)
                        Dim mrRef2 As SolidWorks.Interop.sldworks.IEntity = mr2.ReferenceEntity2(j)
                        If mrRef2 Is Nothing Then
                            Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Information, "Generation Task", taskName, "Reference for Component 2 could not be found", "")
                        Else
                            ' Get the component and select the Persistent ID of the entity
                            Dim corrEnt2 As Entity = swComp2.GetCorresponding(mrRef2)
                            If corrEnt2 IsNot Nothing Then
                                corrEnt2.Select2(True, 1)
                            Else
                            End If
                        End If


                        ' Create the mate
                        ' ToDo: Set the alignment conditions
                        Dim mr1Align = mr1.ReferenceAlignment(j)
                        Dim mr2Align = mr2.ReferenceAlignment(j)
                        Select Case mr1.ReferenceType(j)
                            Case swMateReferenceType_e.swMateReferenceType_Coincident
                                Dim coincMateData As CoincidentMateFeatureData = assy.CreateMateData(swMateType_e.swMateCOINCIDENT)
                                Dim EntitiesToMate(1) As Object
                                EntitiesToMate(0) = SelMgr.GetSelectedObject6(1, 1)
                                EntitiesToMate(1) = SelMgr.GetSelectedObject6(2, 1)
                                coincMateData.EntitiesToMate = EntitiesToMate
                                coincMateData.MateAlignment = GetAlignment(mr1Align, mr2Align)
                                assy.CreateMate(coincMateData)
                            Case swMateReferenceType_e.swMateReferenceType_Concentric
                                Dim concMateData As ConcentricMateFeatureData = assy.CreateMateData(swMateType_e.swMateCONCENTRIC)
                                concMateData.EntitiesToMate(0) = SelMgr.GetSelectedObject6(1, 1)
                                concMateData.EntitiesToMate(1) = SelMgr.GetSelectedObject6(2, 1)
                                concMateData.MateAlignment = GetAlignment(mr1Align, mr2Align)
                                assy.CreateMate(concMateData)
                            Case swMateReferenceType_e.swMateReferenceType_Parallel
                                Dim parMateData As CoincidentMateFeatureData = assy.CreateMateData(swMateType_e.swMatePARALLEL)
                                parMateData.EntitiesToMate(0) = SelMgr.GetSelectedObject6(1, 1)
                                parMateData.EntitiesToMate(1) = SelMgr.GetSelectedObject6(2, 1)
                                parMateData.MateAlignment = GetAlignment(mr1Align, mr2Align)
                                assy.CreateMate(parMateData)
                            Case swMateReferenceType_e.swMateReferenceType_Tangent
                                Dim tanMateData As CoincidentMateFeatureData = assy.CreateMateData(swMateType_e.swMateTANGENT)
                                tanMateData.EntitiesToMate(0) = SelMgr.GetSelectedObject6(1, 1)
                                tanMateData.EntitiesToMate(1) = SelMgr.GetSelectedObject6(2, 1)
                                tanMateData.MateAlignment = GetAlignment(mr1Align, mr2Align)
                                assy.CreateMate(tanMateData)
                            Case Else
                        End Select

                        ' Determine Alignment Condition
                        Dim alignmentCondition As swMateAlign_e = GetAlignment(mr1.ReferenceAlignment(j), mr2.ReferenceAlignment(j))

                        Dim errorStatus As Integer = 0
                        Dim mate As Mate2 = assy.AddMate5(GetMateType(mr1Type), alignmentCondition, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, errorStatus)
                        If errorStatus = SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_NoError Then
                            Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Information, "Generation Task", taskName, "Mate added with no errors", "")
                        Else
                            Select Case errorStatus
                                Case SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_ErrorUknown
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "SOLIDWORKS reported a mate error - Error Unknown", "")
                                Case SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_IncorrectAlignment
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "SOLIDWORKS reported a mate error - Incorrect Alignment", "")
                                Case SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_IncorrectGearRatios
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "SOLIDWORKS reported a mate error - Incorrect Gear Ratios", "")
                                Case SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_IncorrectMateType
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "SOLIDWORKS reported a mate error - Incorrect Mate Type", "")
                                Case SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_IncorrectSelections
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "SOLIDWORKS reported a mate error - Incorrect Selections", "")
                                Case SolidWorks.Interop.swconst.swAddMateError_e.swAddMateError_OverDefinedAssembly
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, "SOLIDWORKS reported a mate error - Overdefined Assembly", "")
                                Case Else
                                    Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Error, "Generation Task", taskName, String.Format("SOLIDWORKS reported a mate error <{0}>", errorStatus), "")
                            End Select
                        End If
                        model.Model.ClearSelection2(True)
                        'Clean Up
                        mate = Nothing
                        mrRef1 = Nothing
                        mrRef2 = Nothing
                    Else
                        Me.Report.WriteEntry(ReportingLevel.Minimal, ReportEntryType.Warning, "Generation Task", taskName, "Mismatched mate types", String.Format("Component 1 <0>: mate reference <1>-<2>: type = <3>; Component <4> type = <5>", comp1Name, mr1.Name, i, mr1Type.ToString, comp2Name, mr2.Name, i, mr2Type.ToString))
                    End If
                Next j
                'Clean Up
                If mr1 IsNot Nothing Then mr1 = Nothing
                If mr2 IsNot Nothing Then mr2 = Nothing
            Next i
            'Clean Up
            If swComp1 IsNot Nothing Then swComp1 = Nothing
            If swComp2 IsNot Nothing Then swComp2 = Nothing
            If assy IsNot Nothing Then assy = Nothing
            If swApp IsNot Nothing Then swApp = Nothing
        End Sub

        Function GetMateRefByName(comp As Component2, name As String) As MateReference
            Dim mrFolderFeat As FeatureFolder
            Dim mrFeat As MateReference
            mrFeat = Nothing
            mrFolderFeat = comp.FeatureByName("MateReferences").GetSpecificFeature2()
            Dim mrFeats As Object
            mrFeats = mrFolderFeat.GetFeatures()
            For i = 0 To (mrFolderFeat.GetFeatureCount - 1)
                mrFeat = mrFeats(i).GetSpecificFeature2()
                Dim refCount As Integer = mrFeat.ReferenceEntityCount
                If mrFeat.Name = name Then
                    Return mrFeat
                End If
            Next i
            Return Nothing
        End Function
        Function GetComponentByName(assy As AssemblyDoc, name As String) As Component2

            Dim nameArray As Object = name.Split("/")
            Dim swComp As Component2 = Nothing
            Dim i As Integer
            For i = 0 To UBound(nameArray)
                Dim swCompFeat As Feature
                If i > 0 Then
                    swCompFeat = swComp.FeatureByName(nameArray(i))
                Else
                    swCompFeat = assy.FeatureByName(nameArray(i))
                End If
                If swCompFeat Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                swComp = swCompFeat.GetSpecificFeature2
            Next i
            Return swComp
        End Function
        Function GetEntityName(entityType As swSelectType_e) As String
            Select Case entityType
                Case swSelectType_e.swSelCOORDSYS
                    Return "COORDSYS"
                Case swSelectType_e.swSelDATUMAXES
                    Return "AXIS"
                Case swSelectType_e.swSelDATUMLINES
                    Return "REFLINE"
                Case swSelectType_e.swSelDATUMPLANES
                    Return "PLANE"
                Case swSelectType_e.swSelDATUMPOINTS
                    Return "DATUMPOINT"
                Case swSelectType_e.swSelEDGES
                    Return "EDGE"
                Case swSelectType_e.swSelFACES
                    Return "FACE"
                Case swSelectType_e.swSelPOINTREFS
                    Return "POINTREF"
                Case swSelectType_e.swSelREFCURVES
                    Return "REFCURVE"
                Case swSelectType_e.swSelREFEDGES
                    Return "REFERENCE-EDGE"
                Case swSelectType_e.swSelREFERENCECURVES
                    Return "REFERENCECURVES"
                Case swSelectType_e.swSelSKETCHPOINTS
                    Return "SKETCHPOINT"
                Case swSelectType_e.swSelSKETCHSEGS
                    Return "SKETCHSEGMENT"
                Case swSelectType_e.swSelVERTICES
                    Return "VERTEX"
                Case Else
                    Return String.Empty
            End Select
        End Function
        Function GetMateType(mateRefType As swMateReferenceType_e) As swMateType_e
            Select Case mateRefType
                Case swMateReferenceType_e.swMateReferenceType_Coincident
                    Return swMateType_e.swMateCOINCIDENT
                Case swMateReferenceType_e.swMateReferenceType_Concentric
                    Return swMateType_e.swMateCONCENTRIC
                Case swMateReferenceType_e.swMateReferenceType_Parallel
                    Return swMateType_e.swMatePARALLEL
                Case swMateReferenceType_e.swMateReferenceType_Tangent
                    Return swMateType_e.swMateTANGENT
                Case swMateReferenceType_e.swMateReferenceType_default
                    Return swMateType_e.swMateUNKNOWN
                Case Else
                    Return swMateType_e.swMateUNKNOWN
            End Select
        End Function
        Function GetAlignment(mr1Align As swMateReferenceAlignment_e, mr2Align As swMateReferenceAlignment_e) As swMateAlign_e
            'AGAINST, NONE and SAME are obsoleted, use only ALIGNED, ANTI_ALIGNED, CLOSEST
            Select Case mr1Align
                Case swMateReferenceAlignment_e.swMateReferenceAlignment_Aligned
                    Select Case mr2Align
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Aligned
                            Return swMateAlign_e.swMateAlignALIGNED
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_AntiAligned
                            Return swMateAlign_e.swMateAlignANTI_ALIGNED
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Closest
                            Return swMateAlign_e.swMateAlignCLOSEST
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Any
                            Return swMateAlign_e.swMateAlignCLOSEST
                        Case Else
                            Return swMateAlign_e.swMateAlignCLOSEST
                    End Select
                Case swMateReferenceAlignment_e.swMateReferenceAlignment_AntiAligned
                    Select Case mr2Align
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Aligned
                            Return swMateAlign_e.swMateAlignANTI_ALIGNED
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_AntiAligned
                            Return swMateAlign_e.swMateAlignANTI_ALIGNED
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Closest
                            Return swMateAlign_e.swMateAlignCLOSEST
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Any
                            Return swMateAlign_e.swMateAlignCLOSEST
                        Case Else
                            Return swMateAlign_e.swMateAlignCLOSEST
                    End Select
                Case swMateReferenceAlignment_e.swMateReferenceAlignment_Closest
                    Return swMateAlign_e.swMateAlignCLOSEST
                Case swMateReferenceAlignment_e.swMateReferenceAlignment_Any
                    Select Case mr2Align
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Aligned
                            Return swMateAlign_e.swMateAlignALIGNED
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_AntiAligned
                            Return swMateAlign_e.swMateAlignANTI_ALIGNED
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Closest
                            Return swMateAlign_e.swMateAlignCLOSEST
                        Case swMateReferenceAlignment_e.swMateReferenceAlignment_Any
                            Return swMateAlign_e.swMateAlignCLOSEST
                        Case Else
                            Return swMateAlign_e.swMateAlignCLOSEST
                    End Select
                Case Else
                    Return swMateAlign_e.swMateAlignCLOSEST
            End Select
        End Function
    End Class

    <GenerationTask("Drive Custom Bend Allowance",
                    "Set custom bend allowance parameters for a single bend feature",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Parts, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
    Public Class DriveCustomBendAllowance
        Inherits GenerationTask

        Private Const FEAT_NAME As String = "GenerationTask.bendFeatureName"
        Private Const BA_TYPE As String = "GenerationTask.bendAllowanceType"
        Private Const BA_VALUE As String = "GenerationTask.bendAllowanceValue"

        Private Enum BendAllowanceType
            BendAllowance = swBendAllowanceTypes_e.swBendAllowanceDirect
            BendDeduction = swBendAllowanceTypes_e.swBendAllowanceDeduction
            BendTable = swBendAllowanceTypes_e.swBendAllowanceBendTable
            BendCalculationTable = swBendAllowanceTypes_e.swBendAllowanceBendCalculationTable
            KFactor = swBendAllowanceTypes_e.swBendAllowanceKFactor
        End Enum

        Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
            Get
                Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(FEAT_NAME, "Bend Feature Name", "Name of the Sheetmetal bend SUBFEATUE (ex. EdgeBend3)", "Features"),
                    New ComponentTaskParameterInfo(BA_TYPE, "Bend Allowance Type", "Type of bend allowance to apply", "Inputs", New ComponentTaskParameterMetaData(DriveWorks.Forms.DataModel.PropertyBehavior.StandardOptionDynamicManual, GetType(BendAllowanceType), BendAllowanceType.KFactor)),
                    New ComponentTaskParameterInfo(BA_VALUE, "Value", "Numeric value for Bend Allowance, Bend Deduction, or K Factor; Name of Bend Table for Bend Table or Calculation Bend Table; Ignored for Non-Custom", "Inputs")
                    }
            End Get
        End Property

        Private Const taskName = "Drive Custom Bend Allowance"

        Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
            Dim featName As String = String.Empty
            Dim baType As String = String.Empty
            Dim baValue As String = String.Empty
            Dim baDoubleValue As Double

            ' Verify that we got data in all of the inputs
            If Not Me.Data.TryGetParameterValue(FEAT_NAME, featName) Then
                ' No entity was received to offset, can't go on
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "Feature name not specified", "No value was received for the bend feature name")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(BA_TYPE, baType) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "No bend allowance type provided", "No value provided for the bend allowance type")
                Return
            ElseIf Not Me.Data.TryGetParameterValue(BA_VALUE, baValue) Then
                ' No entity was received to offset, can't go on
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "Value not specified", "No value was received for the bend allowance value")
                Return
            ElseIf Not Double.TryParse(baValue, baDoubleValue) Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Generation Task", taskName, "Unable to convert value to a number", baValue)
                Return
            End If

            Dim swApp As SldWorks = model.Application.Instance
            Dim Part As ModelDoc2 = model.Model
            Dim boolstatus As Boolean = Part.Extension.SelectByID2(featName, "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Dim SelMgr As SelectionMgr = Part.SelectionManager
            Dim feat As Feature = SelMgr.GetSelectedObject6(1, -1)
            If feat IsNot Nothing Then
                Dim smFeat As SheetMetalFeatureData
                smFeat = feat.GetDefinition
                If smFeat IsNot Nothing Then
                    Dim custBend As CustomBendAllowance
                    custBend = Part.FeatureManager.CreateCustomBendAllowance()
                    Select Case baType
                        Case "Non-Custom"
                            smFeat.SetCustomBendAllowance(Nothing)
                        Case "BendCalculationTable"
                            custBend.Type = swBendAllowanceTypes_e.swBendAllowanceBendCalculationTable
                        Case "BendTable"
                            custBend.Type = swBendAllowanceTypes_e.swBendAllowanceBendTable
                        Case "Deduction"
                            custBend.Type = swBendAllowanceTypes_e.swBendAllowanceDeduction
                            custBend.BendDeduction = baDoubleValue
                        Case "Direct"
                            custBend.Type = swBendAllowanceTypes_e.swBendAllowanceDirect
                            custBend.BendAllowance = baDoubleValue
                        Case "KFactor"
                            custBend.Type = swBendAllowanceTypes_e.swBendAllowanceKFactor
                            custBend.KFactor = baDoubleValue
                    End Select

                    custBend.Type = swBendAllowanceTypes_e.swBendAllowanceKFactor

                    smFeat.SetCustomBendAllowance(custBend)
                End If
            End If

        End Sub
    End Class





End Namespace