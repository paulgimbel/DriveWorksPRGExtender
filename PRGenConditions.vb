Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Policy
Imports DriveWorks
Imports DriveWorks.Components
Imports DriveWorks.Components.Tasks
Imports DriveWorks.Forms.DataModel
Imports DriveWorks.Reporting
Imports DriveWorks.SolidWorks
Imports DriveWorks.SolidWorks.Generation
Imports DriveWorks.SolidWorks.Generation.Proxies
Imports DriveWorks.Specification.StandardTasks
Imports Microsoft.VisualBasic.ApplicationServices
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

<GenerationTaskCondition("Has Rebuild Errors",
                         "Checks for rebuild errors in your model or drawing",
                         "embedded://DriveWorksPRGExtender.bomb.bmp",
                         "PRG Tools",
                         GenerationTaskScope.All,
                         ComponentTaskSequenceLocation.After Or ComponentTaskSequenceLocation.PreClose Or ComponentTaskSequenceLocation.Before)>
Public Class HasError
    Inherits GenerationTaskCondition

    Private Const INCLUDE_WARNINGS As String = "IncludeWarnings"
    'Private Const INCLUDE_BODYCHECK As String = "IncludeBodyCheck"  ' Enables VERIFICATION ON REBUILD/ADVANCED BODY CHECKING

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(INCLUDE_WARNINGS,
                                               "Include Warnings",
                                               "Will return TRUE if the What's Wrong dialog returns rebuild warnings or errors.",
                                               "PRG Tools",
                                               New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), False))
                                                                                         }
        End Get
    End Property

    Protected Overrides Function Evaluate(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings) As Boolean
        Dim includeWarnings As Boolean = False
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_WARNINGS, includeWarnings) Then
            Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Condition", "Has Warnings/Errors", "Input parameter needs to be True or False", "Default value of FALSE will be used.")
        End If

        Dim featureArray As Object = Nothing
        Dim codeArray As Object = Nothing
        Dim warningsArray As Object = Nothing
        Dim returnStatus As Boolean = model.Model.Extension.GetWhatsWrong(featureArray, codeArray, warningsArray)
        If featureArray Is Nothing Then
            Return False
        Else
            ' If there are errors, we will report them back to the Model Gen report in Verbose Mode
            Dim messageText As String
            Dim hasError As Boolean = False
            Dim hasWarning As Boolean = False
            For i = 0 To UBound(featureArray)
                Dim isWarning As Boolean
                If Boolean.TryParse(warningsArray(i), isWarning) Then
                    If isWarning Then hasWarning = True Else hasError = True
                    messageText = String.Format("({0}) Feature: {1}; Error: ", IIf(isWarning, "Warning", "Error"), featureArray(i))
                    Select Case codeArray(i)
                        Case swFeatureError_e.swFeatureErrorExtrusionBadGeometricConditions
                            String.Format("{0} Bad Geometric Conditions", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionBadGeometricConditions
                            String.Format("{0} Unable To create this extruded feature due To geometric conditions", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionBossContourInvalid
                            String.Format("{0} Bosses require one Or more closed contours that Do Not self-intersect", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionBossContourOpenAndClosed
                            String.Format("{0} Bosses cannot have both open And closed contours", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionCutContourInvalid
                            String.Format("{0} Extruded cuts require at least one closed Or open contour that does Not self-intersect", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionCutContourOpenAndClosed
                            String.Format("{0} Extruded cuts cannot have both open And closed contours", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionDisjoint
                            String.Format("{0} Feature would create a disjoint body; direction may be wrong", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionNoEndFound
                            String.Format("{0} Cannot locate End Of feature", messageText)
                        Case swFeatureError_e.swFeatureErrorExtrusionOpenCutContourInvalid
                            String.Format("{0} Open extruded cuts require a Single open contour that does Not self-intersect", messageText)
                        Case swFeatureError_e.swFeatureErrorFeatureDeprecated
                            String.Format("{0} Feature has been depricated", messageText)
                        Case swFeatureError_e.swFeatureErrorFeatureObsolete
                            String.Format("{0} Feature is obsolete", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletCannotExtend
                            String.Format("{0} Selected elements cannot be extended To intersect", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletInvalidRadius
                            String.Format("{0} Invalid fillet radius Or a face blend fillet recommended", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletModelGeometry
                            String.Format("{0} Failed To create fillet due To model geometry", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletNoEdge
                            String.Format("{0} Edge For fillet/chamfer does not exist", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletNoFace
                            String.Format("{0} Face For fillet/chamfer does not exist", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletNoLoop
                            String.Format("{0} Loop For fillet/chamfer does not exist", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletRadiusEliminateElement
                            String.Format("{0} Specified radius would eliminate one Of the elements", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletRadiusTooBig
                            String.Format("{0} Radius Is too big Or the elements are tangent Or nearly tangent", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletRadiusTooBig2
                            String.Format("{0} Fillet radius is too long", messageText)
                        Case swFeatureError_e.swFeatureErrorFilletRadiusTooSmall
                            String.Format("{0} Fillet radius is too small", messageText)
                        Case swFeatureError_e.swFeatureErrorMateBroken
                            String.Format("{0} One Or more mate entities were suppressed", messageText)
                        Case swFeatureError_e.swFeatureErrorMateDanglingGeometry
                            String.Format("{0} Mate points To dangling geometry", messageText)
                        Case swFeatureError_e.swFeatureErrorMateEntityFailed
                            String.Format("{0} Mating Is Not supported For one Of the components Or one Of the components cannot currently be modified", messageText)
                        Case swFeatureError_e.swFeatureErrorMateEntityNotLinear
                            String.Format("{0} Non - linear edges cannot be used For mating", messageText)
                        Case swFeatureError_e.swFeatureErrorMateFailedCreatingSurface
                            String.Format("{0} Mating surface type Is not supported", messageText)
                        Case swFeatureError_e.swFeatureErrorMateIlldefined
                            String.Format("{0} This mate cannot be solved.", messageText)
                        Case swFeatureError_e.swFeatureErrorMateInvalidEdge
                            String.Format("{0} One Of the edges Of this mate Is suppressed, invalid, Or no longer present", messageText)
                        Case swFeatureError_e.swFeatureErrorMateInvalidEntity
                            String.Format("{0} One Of the entities Of this mate Is suppressed, invalid, Or no longer present", messageText)
                        Case swFeatureError_e.swFeatureErrorMateInvalidFace
                            String.Format("{0} One Of the faces Of this mate Is suppressed, invalid, Or no longer present", messageText)
                        Case swFeatureError_e.swFeatureErrorMateOverdefined
                            String.Format("{0} This mate Is over-defining the assembly; consider deleting some Of the over-defining mates", messageText)
                        Case swFeatureError_e.swFeatureErrorMateUnknownTangent
                            String.Format("{0} Tangent Not satisfied", messageText)
                        Case swFeatureError_e.swFeatureErrorNone
                            String.Format("{0} No Error", messageText)
                        Case swFeatureError_e.swFeatureErrorUnknown
                            String.Format("{0} Unknown Error", messageText)
                        Case swFeatureError_e.swSketchErrorExtRefFail
                            String.Format("{0} Sketch Error", messageText)
                        Case Else
                            String.Format("{0} Unknown Error", messageText)
                    End Select
                    Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Condition", "Has Errors/Warnings", messageText, "")
                End If
            Next i
            If hasError Or (includeWarnings And hasWarning) Then
                Return True
            Else
                Return False
            End If
        End If

    End Function
End Class

<GenerationTaskCondition("Needs Rebuild",
                         "Indicates whether SOLIDWORKS thinks your model or drawing needs to be rebuilt",
                         "embedded://DriveWorksPRGExtender.bomb.bmp",
                         "PRG Tools",
                         GenerationTaskScope.All,
                         ComponentTaskSequenceLocation.After Or ComponentTaskSequenceLocation.PreClose Or ComponentTaskSequenceLocation.Before)>
Public Class NeedsRebuild
    Inherits GenerationTaskCondition

    Protected Overrides Function Evaluate(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings) As Boolean
        Dim needsRebuild As swModelRebuildStatus_e = model.Model.Extension.NeedsRebuild2()
        Select Case needsRebuild
            Case swModelRebuildStatus_e.swModelRebuildStatus_FrozenFeatureNeedsRebuild
                Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Condition", "Needs Rebuild", "SOLIDWORKS model is frozen and requires a rebuild.", "")
                Return True
            Case swModelRebuildStatus_e.swModelRebuildStatus_NonFrozenFeatureNeedsRebuild
                Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Condition", "Needs Rebuild", "SOLIDWORKS model is not frozen and requires a rebuild.", "")
                Return True
            Case swModelRebuildStatus_e.swModelRebuildStatus_FullyRebuilt
                Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Condition", "Needs Rebuild", "SOLIDWORKS model is fully rebuilt.", "")
                Return False
            Case Else
                Return False  ' Not sure what the default should be for this case
        End Select
    End Function
End Class

