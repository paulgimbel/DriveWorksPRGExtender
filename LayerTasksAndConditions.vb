Imports System.Drawing
Imports System.IO
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml.Serialization
Imports DriveWorks.Components
Imports DriveWorks.Components.Tasks
Imports DriveWorks.Forms.DataModel
Imports DriveWorks.Reporting
Imports DriveWorks.SolidWorks
Imports DriveWorks.SolidWorks.Generation
Imports DriveWorks.SolidWorks.Generation.Proxies
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

<GenerationTask("Delete Drawing Layer",
                    "Deletes a layer from a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class DeleteLayer
    Inherits GenerationTask

    Private Const LAYER_NAME As String = "LayerName"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAME, "Layer Name", "Name of the layer to be deleted", "Inputs")
                    }
        End Get
    End Property

    Private Const TASKNAME = "Delete Layer"

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim layerName As String = String.Empty
        Dim layerNameList As New List(Of String)
        If Not Me.Data.TryGetParameterValue(LAYER_NAME, layerName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Layer name parameter could not be read")
        Else
            layerNameList = layerName.Split("|").ToList()
        End If

        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            layerMgr = model.Model.GetLayerManager
            For Each name In layerNameList
                layer = layerMgr.GetLayer(name)
                If layer Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "A layer with that name could not be found.")
                Else
                    If layerMgr.DeleteLayer(layerName) Then
                        Me.SetExecutionResult(TaskExecutionResult.Success, "Layer deleted successfully.")
                    Else
                        Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to delete layer.")
                    End If
                End If
            Next
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try
        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", taskName))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub
End Class

<GenerationTask("Activate Drawing Layer",
                    "Changes the active layer in a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class SetActiveLayer
    Inherits GenerationTask

    Private Const LAYER_NAME As String = "LayerName"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAME, "Layer Name", "Name of the layer to be activated", "Inputs")
                    }
        End Get
    End Property

    Private Const TASKNAME = "Activate Drawing Layer"

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim layerName As String = String.Empty
        If Not Me.Data.TryGetParameterValue(LAYER_NAME, layerName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Layer name parameter could not be read")
        End If

        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            layerMgr = model.Model.GetLayerManager
            layer = layerMgr.GetLayer(layerName)
            If layer Is Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "A layer with that name could not be found.")
            Else
                If layerMgr.SetCurrentLayer(layerName) = 1 Then
                    Me.SetExecutionResult(TaskExecutionResult.Success, "Layer activated successfully.")
                Else
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to activate layer.")
                End If
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try
        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", taskName))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub
End Class

<GenerationTask("Add Drawing Layer",
                    "Adds a new layer to a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class AddLayer
    Inherits GenerationTask

    Private Const LAYER_NAME As String = "LayerName"
    Private Const LAYER_DESC As String = "LayerDescription"
    Private Const LAYER_COLOR As String = "LayerColor"
    Private Const LAYER_STYLE As String = "LayerStyle"
    Private Const LAYER_WIDTH As String = "LayerWidth"
    Private Const LAYER_PRINT As String = "LayerPrint"

    Private Const TASKNAME = "Add Layer"

    Private Enum LineStyles
        DEFAULTSTYLE = swLineStyles_e.swLineDEFAULT
        CONTINUOUS = swLineStyles_e.swLineCONTINUOUS
        HIDDEN = swLineStyles_e.swLineHIDDEN
        PHANTOM = swLineStyles_e.swLinePHANTOM
        CHAIN = swLineStyles_e.swLineCHAIN
        CENTER = swLineStyles_e.swLineCENTER
        STITCH = swLineStyles_e.swLineSTITCH
        CHAINTHICK = swLineStyles_e.swLineCHAINTHICK
    End Enum

    Private Enum LineWeights
        NONE = swLineWeights_e.swLW_NONE
        THIN = swLineWeights_e.swLW_THIN
        NORMAL = swLineWeights_e.swLW_NORMAL
        THICK = swLineWeights_e.swLW_THICK
        THICK2 = swLineWeights_e.swLW_THICK2
        THICK3 = swLineWeights_e.swLW_THICK3
        THICK4 = swLineWeights_e.swLW_THICK4
        THICK5 = swLineWeights_e.swLW_THICK5
        THICK6 = swLineWeights_e.swLW_THICK6
        NUMBER = swLineWeights_e.swLW_NUMBER
        LAYER = swLineWeights_e.swLW_LAYER
        CUSTOM = swLineWeights_e.swLW_CUSTOM
    End Enum

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAME,
                                                   "Layer Name",
                                                   "Name of the layer to be added",
                                                   "Inputs",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(LAYER_DESC,
                                                   "Layer Description",
                                                   "Description to be added to the layer",
                                                   "Settings",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(LAYER_COLOR,
                                                   "Line Color",
                                                   "Color to make the entities on the layer (in Red|Green|Blue, ex. 255|0|0 for red)",
                                                   "Settings",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Color), Color.Black)),
                    New ComponentTaskParameterInfo(LAYER_STYLE,
                                                   "Line Style",
                                                   "Line style for the geometry on the layer",
                                                   "Settings",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual,
                                                                                      GetType(LineStyles),
                                                                                      LineStyles.DEFAULTSTYLE)),
                    New ComponentTaskParameterInfo(LAYER_WIDTH,
                                                   "Line Width",
                                                   "Line width for the geometry on the layer",
                                                   "Settings",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual,
                                                                                      GetType(LineWeights),
                                                                                      LineWeights.NORMAL)),
                    New ComponentTaskParameterInfo(LAYER_PRINT,
                                                   "Layer Printable",
                                                   "Determines whether the layer will be included in prints",
                                                   "Settings",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True))
                    }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim layerName As String = String.Empty
        Dim layerDesc As String = String.Empty
        Dim layerColorString As String = String.Empty
        Dim layerColorList As New List(Of Integer)
        Dim layerColorInt As Integer
        Dim layerStyle As swLineStyles_e
        Dim layerWidth As swLineWeights_e
        Dim layerPrintable As Boolean

        ' Validate the inputs
        If Not Me.Data.TryGetParameterValue(LAYER_NAME, layerName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Layer name parameter could not be read")
            Return
        End If
        If Not Me.Data.TryGetParameterValue(LAYER_DESC, layerDesc) Then warnings.Add("Unable to read description input")
        If Not Me.Data.TryGetParameterValue(LAYER_COLOR, layerColorString) Then
            warnings.Add("Unable to read color input")
        Else
            If layerColorString.Contains("|") Then
                For Each val As String In layerColorString.Split("|")
                    Dim colorVal As Integer
                    If Integer.TryParse(val, colorVal) Then
                        layerColorList.Add(colorVal)
                    Else
                        warnings.Add(String.Format("Unable to convert the layer color string <{0}> to a series of three integer values [error processing <{1}>]", layerColorString, val))
                        layerColorList.Add(0)
                    End If
                Next val
                If layerColorList.Count <> 3 Then
                    warnings.Add(String.Format("The color string needs to be three integer values, defaulting to 255|255|255 instead of <{0}>", layerColorString))
                Else
                    layerColorInt = ColorTranslator.ToWin32(Color.FromArgb(layerColorList(0), layerColorList(1), layerColorList(2)))
                End If
            Else
                ' Try to convert the color from a known color
                Try
                    layerColorInt = ColorTranslator.ToWin32(Color.FromName(layerColorString))
                Catch ex As Exception
                    warnings.Add("Unable to convert from known color, defaulting to black")
                    layerColorInt = ColorTranslator.ToWin32(Color.Black)
                End Try
            End If
        End If
        If Not Me.Data.TryGetParameterValue(Of LineStyles)(LAYER_STYLE, layerStyle) Then
            warnings.Add("Unable to read line style, defaulting to DEFAULT")
            layerStyle = swLineStyles_e.swLineDEFAULT
        End If
        If Not Me.Data.TryGetParameterValue(Of LineWeights)(LAYER_WIDTH, layerWidth) Then
            warnings.Add("Unable to read line weight, defaulting to NORMAL")
            layerWidth = swLineWeights_e.swLW_NORMAL
        End If
        If Not Me.Data.TryGetParameterValueAsBoolean(LAYER_PRINT, layerPrintable) Then
            warnings.Add("Unable to read layer printable value, defaulting to TRUE")
            layerPrintable = True
        End If

        'Declare the objects that we are going to have to retrieve (and release)
        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            layerMgr = model.Model.GetLayerManager
            layer = layerMgr.GetLayer(layerName)
            If layer IsNot Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "A layer with that name already exists.")
            Else
                Dim result As Integer = layerMgr.AddLayer(layerName, layerDesc, layerColorInt, layerStyle, layerWidth)  ' The result is not an enum, it's just "1 if the layer was created successfully"
                If result = 1 Then
                    Me.SetExecutionResult(TaskExecutionResult.Success, "Layer added successfully.")
                    ' Setting the printability, as that cannot be set with the constructor...I know, right?
                    layer = layerMgr.GetLayer(layerName)
                    If layer IsNot Nothing Then
                        layer.Printable = layerPrintable
                    Else
                        warnings.Add("Unable to retrieve the newly created layer to set the printable setting")
                    End If
                Else
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to add layer.")
                End If
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If

    End Sub
End Class

<GenerationTask("Move Selected Items To New Drawing Layer",
                    "Moves the currently selected items in a drawing to a specified layer",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class MoveToLayer
    Inherits GenerationTask

    Private Const LAYER_NAME As String = "LayerName"

    Private Const TASKNAME As String = "Move Selected Items to Layer"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAME, "Layer Name", "Name of the target layer", "Inputs")
                    }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim layerName As String = String.Empty
        If Not Me.Data.TryGetParameterValue(LAYER_NAME, layerName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Layer name parameter could not be read")
        End If

        Dim selMgr As SelectionMgr = Nothing
        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            selMgr = model.Model.SelectionManager
            If selMgr Is Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to obtain SOLIDWORKS selection manager.")
                Return
            End If
            For item = 1 To selMgr.GetSelectedObjectCount2(-1)
                Dim currentItem As Object = selMgr.GetSelectedObject6(item, -1)
                Select Case currentItem.GetType
                    Case GetType(DisplayDimension)
                        Dim dimItem As DisplayDimension = TryCast(currentItem, DisplayDimension)
                        If dimItem IsNot Nothing Then
                            Dim annotItem As Annotation = dimItem.GetAnnotation()
                            If annotItem IsNot Nothing Then
                                annotItem.Layer = layerName
                            Else
                                warnings.Add("Unable to get annotation object for display dimension")
                            End If
                        Else
                            warnings.Add("Unable to get display dimension object")
                        End If
                    Case GetType(Note)
                        Dim noteItem As Note = TryCast(currentItem, Note)
                        If noteItem IsNot Nothing Then
                            Dim annotItem As Annotation = noteItem.GetAnnotation()
                            If annotItem IsNot Nothing Then
                                annotItem.Layer = layerName
                            Else
                                warnings.Add("Unable to get annotation object for note")
                            End If
                        Else
                            warnings.Add("Unable to get note object")
                        End If
                    Case GetType(SketchSegment)
                        Dim SketchSeg As SketchSegment = TryCast(currentItem, SketchSegment)
                        If SketchSeg IsNot Nothing Then
                            SketchSeg.Layer = layerName
                        Else
                            warnings.Add("Unable to get sketch segment object")
                        End If
                    Case GetType(SketchPoint)
                        Dim SketchPt As SketchPoint = TryCast(currentItem, SketchPoint)
                        If SketchPt IsNot Nothing Then
                            SketchPt.Layer = layerName
                        Else
                            warnings.Add("Unable to get sketch point object")
                        End If
                    Case Else
                        Try
                            currentItem.Layer = layerName
                        Catch ex As Exception
                            warnings.Add("Unable to get object type")
                        End Try
                End Select
            Next item
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub
End Class

<GenerationTask("Drive Drawing Layer Visibility",
                    "Shows or hides specified layers in a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class DriveLayerVisibility
    Inherits GenerationTask

    Private Const LAYER_NAMES As String = "LayerNames"
    Private Const LAYER_STATE As String = "LayerState"

    Private Const TASKNAME = "Drive Drawing Layer Visibility"

    Private Enum VisibleState
        Hide = 0
        Show = 1
    End Enum

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAMES, "Layer Names", "Pipe-Delimited list of layers to show or hide", "Inputs"),
                    New ComponentTaskParameterInfo(LAYER_STATE, "Layer State", "'Show' to show the specified layers or 'Hide' to hide them.", "Inputs", New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicDefault, GetType(VisibleState), VisibleState.Show))
                    }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)

        ' Validate Inputs
        Dim layerNames As String = String.Empty
        ' Make sure we have a layer name.  If not, quit
        If Not Me.Data.TryGetParameterValue(LAYER_NAMES, True, layerNames) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No layer name given.")
            Return
        Else
            If String.IsNullOrEmpty(layerNames) Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No layer name given.")
                Return
            End If
        End If
        ' Check to see that a valid Visibility State was provided
        Dim layerNameList = layerNames.Split("|").ToList
        Dim layerState As VisibleState
        If Not Me.Data.TryGetParameterValue(Of VisibleState)(LAYER_STATE, layerState) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No layer state given.")
            Return
        End If

        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            For Each layerName In layerNameList
                layerMgr = model.Model.GetLayerManager
                layer = layerMgr.GetLayer(layerName)
                If layer Is Nothing Then
                    warnings.Add(String.Format("Unable to find layer <{0}>", layerName))
                Else
                    layer.Visible = layerState.Equals(VisibleState.Show)
                End If
            Next
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub
End Class

<GenerationTask("Change Drawing Layer Properties",
                    "Change the name, description, color, line weight or line style of a layer in a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
"PRG Tasks",
                    GenerationTaskScope.Drawings,
                    ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class ChangeLayerProps
    Inherits GenerationTask

    Private Const LAYER_NAME As String = "LayerName"
    Private Const LAYER_NEW_NAME As String = "LayerNewName"
    Private Const LAYER_DESC As String = "LayerDescription"
    Private Const LAYER_COLOR As String = "LayerColor"
    Private Const LAYER_STYLE As String = "LayerStyle"
    Private Const LAYER_WEIGHT As String = "LayerWidth"
    Private Const LAYER_PRINT As String = "LayerPrint"

    Private Const TASKNAME = "Change Layer Properties"
    Private Enum LineStyles
        DEFAULTSTYLE = swLineStyles_e.swLineDEFAULT
        CONTINUOUS = swLineStyles_e.swLineCONTINUOUS
        HIDDEN = swLineStyles_e.swLineHIDDEN
        PHANTOM = swLineStyles_e.swLinePHANTOM
        CHAIN = swLineStyles_e.swLineCHAIN
        CENTER = swLineStyles_e.swLineCENTER
        STITCH = swLineStyles_e.swLineSTITCH
        CHAINTHICK = swLineStyles_e.swLineCHAINTHICK
    End Enum

    Private Enum LineWeights
        NONE = swLineWeights_e.swLW_NONE
        THIN = swLineWeights_e.swLW_THIN
        NORMAL = swLineWeights_e.swLW_NORMAL
        THICK = swLineWeights_e.swLW_THICK
        THICK2 = swLineWeights_e.swLW_THICK2
        THICK3 = swLineWeights_e.swLW_THICK3
        THICK4 = swLineWeights_e.swLW_THICK4
        THICK5 = swLineWeights_e.swLW_THICK5
        THICK6 = swLineWeights_e.swLW_THICK6
        NUMBER = swLineWeights_e.swLW_NUMBER
        LAYER = swLineWeights_e.swLW_LAYER
        CUSTOM = swLineWeights_e.swLW_CUSTOM
    End Enum
    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAME,
                                                   "Layer Name",
                                                   "Name of the layer to be added",
                                                   "Inputs",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(LAYER_NEW_NAME,
                                                   "New Layer Name",
                                                   "New name of the layer (leave blank to keep current value)",
                                                   "Inputs",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(LAYER_DESC,
                                                   "Layer Description",
                                                   "Description to be added to the layer (leave blank to keep current value)",
                                                   "Settings",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(LAYER_COLOR,
                                                   "Line Color",
                                                   "Color to make the entities on the layer (named colors or Red|Green|Blue, ex. 'Red' or 255|0|0 for red)",
                                                   "Settings",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Color), "Black")),
                    New ComponentTaskParameterInfo(LAYER_STYLE,
                                                   "Line Style",
                                                   "Line style for the geometry on the layer (leave blank to keep current value)",
                                                   "Settings",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual,
                                                                                      GetType(LineStyles), LineStyles.DEFAULTSTYLE)),
                    New ComponentTaskParameterInfo(LAYER_WEIGHT,
                                                   "Line Width",
                                                   "Line width for the geometry on the layer (leave blank to keep current value)",
                                                   "Settings",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual,
                                                                                      GetType(LineWeights), LineWeights.NORMAL)),
                    New ComponentTaskParameterInfo(LAYER_PRINT,
                                                   "Layer Will Print",
                                                   "Determines whether the layer will be included in prints",
                                                   "Settings",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual,
                                                                                      GetType(Boolean), True))
                    }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim layerName As String = String.Empty
        Dim layerNewName As String = String.Empty
        Dim layerDesc As String = String.Empty
        Dim driveStyle As Boolean = False
        Dim driveWeight As Boolean = False
        Dim driveColor As Boolean = False
        Dim drivePrint As Boolean = False
        Dim layerColorString As String = String.Empty
        Dim layerColorList As New List(Of Integer)
        Dim layerColorInt As Integer
        Dim layerStyle As swLineStyles_e
        Dim layerWeight As swLineWeights_e
        Dim layerPrintable As Boolean = True

        ' Validate the inputs
        If Not Me.Data.TryGetParameterValue(LAYER_NAME, layerName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Layer name parameter could not be read")
            Return
        End If
        If Not Me.Data.TryGetParameterValue(LAYER_NEW_NAME, layerNewName) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "No new layer name provided", "Existing value will be used")
        End If
        If Not Me.Data.TryGetParameterValue(LAYER_DESC, layerDesc) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "No new layer description provided", "Existing value will be used")
        End If
        If Not Me.Data.TryGetParameterValue(LAYER_COLOR, layerColorString) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "No new layer color provided", "Existing value will be used")
        Else
            If layerColorString.Contains("|") Then
                For Each val As String In layerColorString.Split("|")
                    Dim colorVal As Integer
                    If Integer.TryParse(val, colorVal) Then
                        layerColorList.Add(colorVal)
                    Else
                        warnings.Add(String.Format("Unable to convert the layer color string <{0}> to a series of three integer values [error processing <{1}>]", layerColorString, val))
                        layerColorList.Add(0)
                    End If
                Next val
                If layerColorList.Count <> 3 Then
                    warnings.Add(String.Format("The color string needs to be three integer values, existing value will be used instead of <{0}>", layerColorString))
                    driveColor = False
                Else
                    layerColorInt = ColorTranslator.ToWin32(Color.FromArgb(layerColorList(0), layerColorList(1), layerColorList(2)))
                    driveColor = True
                End If
            ElseIf layerColorString.StartsWith("#") Then
                ' Try to convert the color from hex
                Try
                    layerColorInt = ColorTranslator.ToWin32(ColorTranslator.FromHtml(layerColorString))
                    driveColor = True
                Catch ex As Exception
                    warnings.Add("Unable to convert from known color, defaulting to black")
                    layerColorInt = ColorTranslator.ToWin32(Color.Black)
                    driveColor = False
                End Try
            Else
                ' Try to convert the color from a known color
                Try
                    layerColorInt = ColorTranslator.ToWin32(Color.FromName(layerColorString))
                    driveColor = True
                Catch ex As Exception
                    warnings.Add("Unable to convert from known color, defaulting to black")
                    layerColorInt = ColorTranslator.ToWin32(Color.Black)
                    driveColor = False
                End Try
            End If
        End If
        If Not Me.Data.TryGetParameterValue(Of swLineStyles_e)(LAYER_STYLE, layerStyle) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "No new layer style provided", "Existing value will be used")
            driveStyle = False
        Else
            driveStyle = True
        End If
        If Not Me.Data.TryGetParameterValue(Of swLineWeights_e)(LAYER_WEIGHT, layerWeight) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "No new layer weight provided", "Existing value will be used")
            driveWeight = False
        Else
            driveWeight = True
        End If
        If Not Me.Data.TryGetParameterValueAsBoolean(LAYER_PRINT, layerPrintable) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "No value provided for layer is printable", "Existing value will be used")
            drivePrint = False
        Else
            drivePrint = True
        End If
        'Declare the objects that we are going to have to retrieve (and release)
        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            layerMgr = model.Model.GetLayerManager
            layer = layerMgr.GetLayer(layerName)
            If layer Is Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "A layer with that name could not be found.")
            Else
                If Not String.IsNullOrEmpty(layerNewName) Then layer.Name = layerNewName
                If Not String.IsNullOrEmpty(layerDesc) Then layer.Description = layerDesc
                If driveStyle Then layer.Style = layerStyle
                If driveColor Then layer.Color = layerColorInt
                If driveWeight Then layer.Width = layerWeight
                If drivePrint Then layer.Printable = layerPrintable
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed successfully", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If

    End Sub
End Class

<GenerationTaskCondition("Has Layers",
                         "Checks for the existence of drawing layers",
                         "embedded://DriveWorksPRGExtender.HasLayer.png",
                         "SOLIDWORKS PowerPack Drawing",
                         GenerationTaskScope.Drawings,
                         ComponentTaskSequenceLocation.After Or ComponentTaskSequenceLocation.PreClose Or ComponentTaskSequenceLocation.Before)>
Public Class HasLayer
    Inherits GenerationTaskCondition

    Private Const LAYER_NAMES As String = "LayerNames"
    Private Const REQUIRE_ALL As String = "RequireAll"  ' Choose between Logical AND and OR

    Private Const TASKNAME = "Has Layers"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(LAYER_NAMES,
                                               "Layer Names",
                                               "Pipe-delimited list of layer names to check",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(REQUIRE_ALL,
                                               "Require All",
                                               "True to require all layers to exist (logical AND), False to require any layer to exist (logical OR), Default = True",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True))
                                                }
        End Get
    End Property

    Protected Overrides Function Evaluate(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings) As Boolean
        ' Validate inputs
        Dim layerNamesString As String = String.Empty
        Dim layerNamesList As New List(Of String)
        Dim requireAll As Boolean = True
        If Not Me.Data.TryGetParameterValue(LAYER_NAMES, True, layerNamesString) Then
            Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Condition", TASKNAME, "Layer Names parameter not provided", "No layers provided, so condition will return True.")
            Return True
        Else
            layerNamesList = layerNamesString.Split("|").ToList
        End If
        If Not Me.Data.TryGetParameterValueAsBoolean(REQUIRE_ALL, requireAll) Then
            Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Condition", TASKNAME, "Require All parameter needs to be True or False", "Default value of FALSE will be used.")
            requireAll = True
        End If


        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Try
            layerMgr = model.Model.GetLayerManager
            If layerMgr Is Nothing Then
                Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Warning, "Model Generation Condition", TASKNAME, "Unable to obtain the SOLIDWORKS Layer Manager for the current drawing", "")
                Return False
            Else
                For Each layerName In layerNamesList
                    layer = layerMgr.GetLayer(layerName)
                    If layer IsNot Nothing Then
                        If Not requireAll Then Return True
                        layer = Nothing
                    Else
                        If requireAll Then Return False
                    End If
                Next
                ' If we got this far, then we've found all of the layers (requireAll = True) or none of the layers (requireAll = False)
                If requireAll Then Return True Else Return False
            End If
            layerMgr = Nothing
        Catch ex As Exception
            Me.Report.WriteEntry(ReportingLevel.Normal, ReportEntryType.Error, "Model Generation Condition", TASKNAME, "Exception Error", ex.Message)
            Return False
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try

    End Function
End Class

<GenerationTask("Move All Entities To Layer",
                    "Move all of the sketch and annotation entities to a different drawing layer",
                    "embedded://DriveWorksPRGExtender.MoveToLayer.png",
                    "SOLIDWORKS PowerPack Drawing",
                    GenerationTaskScope.Drawings,
                    ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class MoveAllOnLayer
    Inherits GenerationTask

    Private Const SOURCE_LAYER_NAME As String = "FromLayerName"
    Private Const TARGET_LAYER_NAME As String = "ToLayerName"
    Private Const INCLUDE_ANN As String = "IncludeAnnotations"
    'Private Const INCLUDE_BLK As String = "IncludeBlocks"           ' Only the first instance of a block is selectable with layer.GetItems()
    Private Const INCLUDE_HAT As String = "IncludeHatches"
    Private Const INCLUDE_SK As String = "IncludeSketchSegments"
    Private Const INCLUDE_PT As String = "IncludeSketchPoints"

    Private Const TASKNAME = "Move All Entities To Layer"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(SOURCE_LAYER_NAME,
                                                   "Source Layer Name",
                                                   "Name of the layer where entities currently reside",
                                                   "Inputs",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(TARGET_LAYER_NAME,
                                                   "Target Layer Name",
                                                   "Name of the target layer for the entities",
                                                   "Inputs",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(INCLUDE_ANN,
                                                   "Include Annotations",
                                                   "Select annotations on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True)),
                    New ComponentTaskParameterInfo(INCLUDE_HAT,
                                                   "Include Hatches",
                                                   "Select hatch patterns on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True)),
                     New ComponentTaskParameterInfo(INCLUDE_SK,
                                                   "Include Sketch Segments",
                                                   "Select sketch segments on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True)),
                     New ComponentTaskParameterInfo(INCLUDE_PT,
                                                   "Include Sketch Points",
                                                   "Select sketch points on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True))
                                                   }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim sourceLayerNameString As String = String.Empty
        Dim sourceLayerNameList As New List(Of String)
        Dim targetLayerName As String = String.Empty
        ' Validate the inputs
        If Not Me.Data.TryGetParameterValue(SOURCE_LAYER_NAME, sourceLayerNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Source Layer Name parameter not provided")
            Return
        Else
            sourceLayerNameList = sourceLayerNameString.Split("|").ToList()
        End If
        If Not Me.Data.TryGetParameterValue(TARGET_LAYER_NAME, True, targetLayerName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Target Layer Name parameter not provided")
            Return
        End If
        Dim includeAnn As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_ANN, includeAnn) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Annotations parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If
        Dim includeHatch As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_HAT, includeHatch) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Hatches parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If
        Dim includeSkSeg As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_SK, includeSkSeg) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Sketch Segments parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If
        Dim includeSkPt As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_PT, includeSkPt) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Sketch Points parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If        'Declare the objects that we are going to have to retrieve (and release)
        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Dim entityCount As Integer = 0
        Try
            layerMgr = model.Model.GetLayerManager
            If layerMgr Is Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to access the SOLIDWORKS Layer Manager for this document.")
                Return
            Else
                For Each layerName In sourceLayerNameList
                    layer = layerMgr.GetLayer(layerName)
                    If layer Is Nothing Then
                        Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, String.Format("A layer with the name <{0}> could not be found.", layerName))
                    Else
                        If includeAnn Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_Annotations)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As Annotation = items(i)
                                    'item.Select(True)
                                    item.Layer = targetLayerName
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                            If includeHatch Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_SketchHatch)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As SketchHatch = items(i)
                                    'item.Select(True)
                                    item.Layer = targetLayerName
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                        If includeSkSeg Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_SketchSegments)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As SketchSegment = items(i)
                                    'item.Select(True)
                                    item.Layer = targetLayerName
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                        If includeSkPt Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_SketchPoint)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As SketchPoint = items(i)
                                    'item.Select(True)
                                    item.Layer = targetLayerName
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} Entities moved successfully", entityCount))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If

    End Sub
End Class

<GenerationTask("Select All Entities On A Layer",
                    "Select all of the sketch and annotation entities on a specific drawing layer",
                    "embedded://DriveWorksPRGExtender.SelectOnLayer.png",
                    "SOLIDWORKS PowerPack Drawing",
                    GenerationTaskScope.Drawings,
                    ComponentTaskSequenceLocation.Before + ComponentTaskSequenceLocation.After)>
Public Class SelectAllOnLayer
    Inherits GenerationTask

    Private Const LAYER_NAME As String = "LayerName"
    Private Const INCLUDE_ANN As String = "IncludeAnnotations"
    'Private Const INCLUDE_BLK As String = "IncludeBlocks"           ' Only the first instance of a block is selectable with layer.GetItems()
    Private Const INCLUDE_HAT As String = "IncludeHatches"
    Private Const INCLUDE_SK As String = "IncludeSketchSegments"
    Private Const INCLUDE_PT As String = "IncludeSketchPoints"

    Private Const TASKNAME = "Select All Entities On Layer"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(LAYER_NAME,
                                                   "Layer Name",
                                                   "Name of the layer where entities currently reside",
                                                   "Inputs",
                                                    New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                    New ComponentTaskParameterInfo(INCLUDE_ANN,
                                                   "Include Annotations",
                                                   "Select annotations on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True)),
                    New ComponentTaskParameterInfo(INCLUDE_HAT,
                                                   "Include Hatches",
                                                   "Select hatch patterns on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True)),
                     New ComponentTaskParameterInfo(INCLUDE_SK,
                                                   "Include Sketch Segments",
                                                   "Select sketch segments on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True)),
                     New ComponentTaskParameterInfo(INCLUDE_PT,
                                                   "Include Sketch Points",
                                                   "Select sketch points on the specified layer",
                                                   "Types",
                                                   New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(Boolean), True))
                                                   }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Dim sourceLayerNameString As String = String.Empty
        Dim sourceLayerNameList As New List(Of String)
        Dim targetLayerName As String = String.Empty
        ' Validate the inputs
        If Not Me.Data.TryGetParameterValue(LAYER_NAME, sourceLayerNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Source Layer Name parameter not provided")
            Return
        Else
            sourceLayerNameList = sourceLayerNameString.Split("|").ToList()
        End If
        Dim includeAnn As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_ANN, includeAnn) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Annotations parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If
        Dim includeHatch As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_HAT, includeHatch) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Hatches parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If
        Dim includeSkSeg As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_SK, includeSkSeg) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Sketch Segments parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If
        Dim includeSkPt As Boolean
        If Not Me.Data.TryGetParameterValueAsBoolean(INCLUDE_PT, includeSkPt) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Include Sketch Points parameter could not be read, please use True or False. DEFAULTING TO TRUE.")
        End If        'Declare the objects that we are going to have to retrieve (and release)
        Dim layerMgr As LayerMgr = Nothing
        Dim layer As Layer = Nothing
        Dim entityCount As Integer = 0
        Try
            layerMgr = model.Model.GetLayerManager
            If layerMgr Is Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to access the SOLIDWORKS Layer Manager for this document.")
                Return
            Else
                For Each layerName In sourceLayerNameList
                    layer = layerMgr.GetLayer(layerName)
                    If layer Is Nothing Then
                        Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, String.Format("A layer with the name <{0}> could not be found.", layerName))
                    Else
                        If includeAnn Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_Annotations)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As Annotation = items(i)
                                    item.Select(True)
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                        If includeHatch Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_SketchHatch)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As SketchHatch = items(i)
                                    item.Select(True)
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                        If includeSkSeg Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_SketchSegments)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As SketchSegment = items(i)
                                    item.Select(True)
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                        If includeSkPt Then
                            Dim items = layer.GetItems(swLayerItemsOption_e.swLayerItemsOption_SketchPoint)
                            If items IsNot Nothing Then
                                For i = 0 To UBound(items)
                                    Dim item As SketchPoint = items(i)
                                    item.Select(True)
                                    entityCount += 1
                                    item = Nothing
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If layer IsNot Nothing Then layer = Nothing
            If layerMgr IsNot Nothing Then layerMgr = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} Entities selected successfully", entityCount))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If

    End Sub
End Class

<GenerationTask("Select Drawing View Entities",
                    "Allows you to select visible named items in a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "PRG Tasks",
                    GenerationTaskScope.Drawings,
                    ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After)>
Public Class SelectDrawingEntities
    Inherits GenerationTask

    Private Const ENTITY_NAMES As String = "EntityNames"
    Private Const VIEW_NAME As String = "ViewName"
    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                    New ComponentTaskParameterInfo(ENTITY_NAMES, "Entity Names", "Pipe-Delimited list of entities to select (ex. FlangeSketch|Line1)", "Inputs"),
                    New ComponentTaskParameterInfo(VIEW_NAME, "View Name", "Name of drawing view to select the entity (leave blank for sheet entities or models)", "Inputs")
                    }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)

        ' Validate Inputs
        Dim entityNames As String = String.Empty
        ' Make sure we have a layer name.  If not, quit
        If Not Me.Data.TryGetParameterValue(ENTITY_NAMES, True, entityNames) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No layer name given.")
            Return
        Else
            If String.IsNullOrEmpty(entityNames) Then
                Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No layer name given.")
                Return
            End If
        End If
        ' Split up the entity names
        Dim entityNameList = entityNames.Split("|").ToList
        Dim viewName As String = String.Empty
        Dim entityInView As Boolean = False
        If Me.Data.TryGetParameterValue(VIEW_NAME, True, viewName) Then entityInView = True

        Dim entity As SolidWorks.Interop.sldworks.Entity = Nothing
        Dim view As View = Nothing
        Dim refModel As ModelDoc2 = Nothing
        Dim refPart As PartDoc = Nothing
        Dim selMgr As SelectionMgr = Nothing
        Dim entityCount As Integer = 0
        Try
            ' In order to make this easier for the user, we're going to try to figure out what type of entity it is.
            ' Our search order will determine preference for identically named items.
            ' Users SHOULD use unique names if they are calling out something by name.
            selMgr = model.Model.SelectionManager
            If selMgr Is Nothing Then warnings.Add("Unable to retrieve the SOLIDWORKS Selection Manager")
            Dim views() As View = model.Drawing.GetViews
            Dim viewNamesList As New List(Of String)
            For viewCounter = 0 To UBound(views)
                viewNamesList.Add(views(viewCounter).Name)
            Next
            If Not (viewNamesList.Contains(viewName)) Then
                Me.SetExecutionResult(TaskExecutionResult.Failed,
                                      String.Format("View with name {0} could not be found.", viewName))
                Return
            End If
            If entityInView Then
                view = model.Drawing.FeatureByName(viewName).GetSpecificFeature()
                If view Is Nothing Then
                    ' Check to see if it's a sheet name
                    Dim sheetNames() As String = model.Drawing.GetSheetNames
                    If sheetNames.Contains(viewName) Then

                    End If
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, String.Format("Unable to retrieve drawing view <{0}>", viewName))
                Else
                    ' Get the entity object by looping through all of the sketch segments
                    Dim viewSketch As Sketch = view.GetSketch()
                    Dim segs As Object = viewSketch.GetSketchSegments()
                    For index = 0 To UBound(segs)
                        Dim seg As SketchSegment = segs(index)
                        If entityNameList.Contains(seg.GetName()) Then
                            view.SelectEntity(seg, True)
                            entityCount += 1
                        End If
                    Next
                    'OLD METHOD - model.Drawing.ActivateView(viewName)
                    'OLD METHOD - For Each itemName In entityNameList
                    'OLD METHOD - TrySelectEntityByName(itemName, model.Model)
                    'OLD METHOD - Next
                End If
            Else
                For Each itemName In entityNameList
                    TrySelectEntityByName(itemName, model.Model)
                Next
                entityCount = selMgr.GetSelectedObjectCount2(-1)
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If entity IsNot Nothing Then entity = Nothing
            If view IsNot Nothing Then view = Nothing
            If refModel IsNot Nothing Then refModel = Nothing
            If refPart IsNot Nothing Then refPart = Nothing
        End Try

        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} entities selected successfully", entityCount))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub

    Private Function TrySelectEntityByName(name As String, model As ModelDoc2) As Boolean
        Dim typeStringList As New Dictionary(Of swSelectType_e, String)
        ' Need to find some way to get these text strings from the swConst interop
        ' This list can probably be pared down to items that we want to select
        typeStringList.Add(swSelectType_e.swSelSKETCHES, "SKETCH")
        typeStringList.Add(swSelectType_e.swSelSKETCHSEGS, "SKETCHSEGMENT")
        typeStringList.Add(swSelectType_e.swSelSKETCHPOINTS, "SKETCHPOINT")
        typeStringList.Add(swSelectType_e.swSelEDGES, "EDGE")
        typeStringList.Add(swSelectType_e.swSelFACES, "FACE")
        typeStringList.Add(swSelectType_e.swSelVERTICES, "VERTEX")
        typeStringList.Add(swSelectType_e.swSelDATUMPLANES, "PLANE")
        typeStringList.Add(swSelectType_e.swSelDATUMAXES, "AXIS")
        typeStringList.Add(swSelectType_e.swSelDATUMPOINTS, "DATUMPOINT")
        typeStringList.Add(swSelectType_e.swSelGTOLS, "GTOL")
        typeStringList.Add(swSelectType_e.swSelDIMENSIONS, "DIMENSION")
        typeStringList.Add(swSelectType_e.swSelNOTES, "NOTE")
        typeStringList.Add(swSelectType_e.swSelCENTERMARKS, "CENTERMARKS")
        typeStringList.Add(swSelectType_e.swSelREFCURVES, "REFCURVE")
        typeStringList.Add(swSelectType_e.swSelSKETCHTEXT, "SKETCHTEXT")
        typeStringList.Add(swSelectType_e.swSelSFSYMBOLS, "SFSYMBOL")
        typeStringList.Add(swSelectType_e.swSelDATUMTAGS, "DATUMTAG")
        typeStringList.Add(swSelectType_e.swSelEXTSKETCHSEGS, "EXTSKETCHSEGMENT")
        typeStringList.Add(swSelectType_e.swSelEXTSKETCHPOINTS, "EXTSKETCHPOINT")
        'typeStringList.Add(swSelectType_e.swSelHELIX, "HELIX")       'swSelectType_e.swSelHELIX = 26, and so does swSelectType_e.swSelREFERENCECURVES
        typeStringList.Add(swSelectType_e.swSelREFERENCECURVES, "REFERENCECURVES")
        typeStringList.Add(swSelectType_e.swSelOLEITEMS, "OLEITEM")
        typeStringList.Add(swSelectType_e.swSelATTRIBUTES, "ATTRIBUTE")
        typeStringList.Add(swSelectType_e.swSelDRAWINGVIEWS, "DRAWINGVIEW")
        typeStringList.Add(swSelectType_e.swSelSECTIONLINES, "SECTIONLINE")
        typeStringList.Add(swSelectType_e.swSelDETAILCIRCLES, "DETAILCIRCLE")
        typeStringList.Add(swSelectType_e.swSelSKETCHHATCH, "SKETCHHATCH")
        typeStringList.Add(swSelectType_e.swSelSECTIONTEXT, "SECTIONTEXT")
        typeStringList.Add(swSelectType_e.swSelCENTERLINES, "CENTERLINE")
        typeStringList.Add(swSelectType_e.swSelHOLETABLEFEATS, "HOLETABLE")
        typeStringList.Add(swSelectType_e.swSelHOLETABLEAXES, "HOLETABLEAXIS")
        typeStringList.Add(swSelectType_e.swSelBOMFEATURES, "BOMFEATURE")
        typeStringList.Add(swSelectType_e.swSelANNOTATIONTABLES, "ANNOTATIONTABLES")
        typeStringList.Add(swSelectType_e.swSelREVISIONTABLE, "REVISIONTABLE")
        typeStringList.Add(swSelectType_e.swSelGENERALTABLEFEAT, "GENERALTABLEFEAT")
        typeStringList.Add(swSelectType_e.swSelPUNCHTABLEFEATS, "PUNCHTABLE")
        typeStringList.Add(swSelectType_e.swSelBLOCKDEF, "BLOCKDEF")
        typeStringList.Add(swSelectType_e.swSelCENTERMARKSYMS, "CENTERMARKSYMS")
        typeStringList.Add(swSelectType_e.swSelCOMPONENTS, "COMPONENT")
        typeStringList.Add(swSelectType_e.swSelMATES, "MATE")
        typeStringList.Add(swSelectType_e.swSelBODYFEATURES, "BODYFEATURE")
        typeStringList.Add(swSelectType_e.swSelREFSURFACES, "REFSURFACE")
        typeStringList.Add(swSelectType_e.swSelINCONTEXTFEAT, "INCONTEXTFEAT")
        typeStringList.Add(swSelectType_e.swSelMATEGROUP, "MATEGROUP")
        typeStringList.Add(swSelectType_e.swSelBREAKLINES, "BREAKLINE")
        typeStringList.Add(swSelectType_e.swSelINCONTEXTFEATS, "INCONTEXTFEATS")
        typeStringList.Add(swSelectType_e.swSelMATEGROUPS, "MATEGROUPS")
        typeStringList.Add(swSelectType_e.swSelCOMPPATTERN, "COMPPATTERN")
        typeStringList.Add(swSelectType_e.swSelWELDS, "WELD")
        typeStringList.Add(swSelectType_e.swSelCTHREADS, "CTHREAD")
        typeStringList.Add(swSelectType_e.swSelDTMTARGS, "DTMTARG")
        typeStringList.Add(swSelectType_e.swSelPOINTREFS, "POINTREF")
        typeStringList.Add(swSelectType_e.swSelDCABINETS, "DCABINET")
        typeStringList.Add(swSelectType_e.swSelEXPLVIEWS, "EXPLODEDVIEWS")
        typeStringList.Add(swSelectType_e.swSelEXPLSTEPS, "EXPLODESTEPS")
        typeStringList.Add(swSelectType_e.swSelEXPLLINES, "EXPLODELINES")
        typeStringList.Add(swSelectType_e.swSelSILHOUETTES, "SILHOUETTE")
        typeStringList.Add(swSelectType_e.swSelCONFIGURATIONS, "CONFIGURATIONS")
        typeStringList.Add(swSelectType_e.swSelARROWS, "VIEWARROW")
        typeStringList.Add(swSelectType_e.swSelZONES, "ZONES")
        typeStringList.Add(swSelectType_e.swSelREFEDGES, "REFERENCE-EDGE")
        typeStringList.Add(swSelectType_e.swSelBOMS, "BOM")
        typeStringList.Add(swSelectType_e.swSelEQNFOLDER, "EQNFOLDER")
        typeStringList.Add(swSelectType_e.swSelIMPORTFOLDER, "IMPORTFOLDER")
        typeStringList.Add(swSelectType_e.swSelVIEWERHYPERLINK, "HYPERLINK")
        typeStringList.Add(swSelectType_e.swSelCOORDSYS, "COORDSYS")
        typeStringList.Add(swSelectType_e.swSelDATUMLINES, "REFLINE")
        typeStringList.Add(swSelectType_e.swSelBOMTEMPS, "BOMTEMP")
        typeStringList.Add(swSelectType_e.swSelSHEETS, "SHEET")
        typeStringList.Add(swSelectType_e.swSelROUTEPOINTS, "ROUTEPOINT")
        typeStringList.Add(swSelectType_e.swSelCONNECTIONPOINTS, "CONNECTIONPOINT")
        typeStringList.Add(swSelectType_e.swSelPOSGROUP, "POSGROUP")
        typeStringList.Add(swSelectType_e.swSelBROWSERITEM, "BROWSERITEM")
        typeStringList.Add(swSelectType_e.swSelFABRICATEDROUTE, "ROUTEFABRICATED")
        typeStringList.Add(swSelectType_e.swSelSKETCHPOINTFEAT, "SKETCHPOINTFEAT")
        typeStringList.Add(swSelectType_e.swSelLIGHTS, "LIGHTS")
        typeStringList.Add(swSelectType_e.swSelSURFACEBODIES, "SURFACEBODY")
        typeStringList.Add(swSelectType_e.swSelSOLIDBODIES, "SOLIDBODY")
        typeStringList.Add(swSelectType_e.swSelFRAMEPOINT, "FRAMEPOINT")
        typeStringList.Add(swSelectType_e.swSelMANIPULATORS, "MANIPULATOR")
        typeStringList.Add(swSelectType_e.swSelPICTUREBODIES, "PICTURE BODY")
        typeStringList.Add(swSelectType_e.swSelLEADERS, "LEADER")
        typeStringList.Add(swSelectType_e.swSelSKETCHBITMAP, "SKETCHBITMAP")
        typeStringList.Add(swSelectType_e.swSelDOWELSYMS, "DOWLELSYM")
        typeStringList.Add(swSelectType_e.swSelEXTSKETCHTEXT, "EXTSKETCHTEXT")
        typeStringList.Add(swSelectType_e.swSelFTRFOLDER, "FTRFOLDER")
        typeStringList.Add(swSelectType_e.swSelSKETCHREGION, "SKETCHREGION")
        typeStringList.Add(swSelectType_e.swSelSKETCHCONTOUR, "SKETCHCONTOUR")
        typeStringList.Add(swSelectType_e.swSelSIMULATION, "SIMULATION")
        typeStringList.Add(swSelectType_e.swSelSIMELEMENT, "SIMULATION_ELEMENT")
        typeStringList.Add(swSelectType_e.swSelWELDMENT, "WELDMENT")
        typeStringList.Add(swSelectType_e.swSelSUBWELDFOLDER, "SUBWELDMENT")
        typeStringList.Add(swSelectType_e.swSelSUBSKETCHINST, "SUBSKETCHINST")
        typeStringList.Add(swSelectType_e.swSelWELDMENTTABLEFEATS, "WELDMENTTABLE")
        typeStringList.Add(swSelectType_e.swSelBODYFOLDER, "BDYFOLDER")
        typeStringList.Add(swSelectType_e.swSelREVISIONTABLEFEAT, "REVISIONTABLEFEAT")
        typeStringList.Add(swSelectType_e.swSelWELDBEADS, "WELDBEADS")
        typeStringList.Add(swSelectType_e.swSelEMBEDLINKDOC, "EMBEDLINKDOC")
        typeStringList.Add(swSelectType_e.swSelJOURNAL, "JOURNAL")
        typeStringList.Add(swSelectType_e.swSelDOCSFOLDER, "DOCSFOLDER")
        typeStringList.Add(swSelectType_e.swSelCOMMENTSFOLDER, "COMMENTSFOLDER")
        typeStringList.Add(swSelectType_e.swSelCOMMENT, "COMMENT")
        typeStringList.Add(swSelectType_e.swSelCAMERAS, "CAMERAS")
        typeStringList.Add(swSelectType_e.swSelMATESUPPLEMENT, "MATESUPPLEMENT")
        typeStringList.Add(swSelectType_e.swSelANNOTATIONVIEW, "ANNVIEW")
        typeStringList.Add(swSelectType_e.swSelSUBSKETCHDEF, "SUBSKETCHDEF")
        typeStringList.Add(swSelectType_e.swSelDISPLAYSTATE, "VISUALSTATE")
        typeStringList.Add(swSelectType_e.swSelTITLEBLOCK, "TITLEBLOCK")
        typeStringList.Add(swSelectType_e.swSelEVERYTHING, "EVERYTHING")
        typeStringList.Add(swSelectType_e.swSelLOCATIONS, "LOCATIONS")
        typeStringList.Add(swSelectType_e.swSelUNSUPPORTED, "UNSUPPORTED")
        typeStringList.Add(swSelectType_e.swSelSWIFTANNOTATIONS, "SWIFTANN")
        typeStringList.Add(swSelectType_e.swSelSWIFTFEATURES, "SWIFTFEATURE")
        typeStringList.Add(swSelectType_e.swSelSWIFTSCHEMA, "SWIFTSCHEMA")
        typeStringList.Add(swSelectType_e.swSelTITLEBLOCKTABLEFEAT, "TITLEBLOCKTABLEFEAT")
        typeStringList.Add(swSelectType_e.swSelOBJGROUP, "OBJGROUP")
        typeStringList.Add(swSelectType_e.swSelCOSMETICWELDS, "COSMETICWELDS")
        typeStringList.Add(swSelectType_e.SwSelMAGNETICLINES, "MAGNETICLINES")
        typeStringList.Add(swSelectType_e.swSelSELECTIONSETFOLDER, "SELECTIONSETFOLDER")
        typeStringList.Add(swSelectType_e.swSelSELECTIONSETNODE, "SUBSELECTIONSETNODE")
        typeStringList.Add(swSelectType_e.swSelHOLESERIES, "HOLESERIES")

        For Each typeString As KeyValuePair(Of swSelectType_e, String) In typeStringList
            If model.Extension.SelectByID2(name, typeString.Value, 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault) Then
                Return True
            End If
        Next
        Return False
    End Function
End Class

<GenerationTask("Convert Drawing View Entities",
                    "Allows you to use Convert Entities to create sketch geometry on the currently selected items in a drawing",
                    "embedded://DriveWorksPRGExtender.bomb.bmp",
                    "SOLIDWORKS PowerPack Drawing",
                    GenerationTaskScope.Drawings,
                    ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After)>
Public Class ConvertDrawingEntities
    Inherits GenerationTask

    Private Const TASKNAME As String = "Convert Drawing View Entities"
    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {}
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        Try
            If model.Model.SketchManager.SketchUseEdge3(False, False) Then
            Else
                Me.SetExecutionResult(TaskExecutionResult.Failed, "Unable to convert selected entities")
                Return
            End If
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
        End Try
        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub

End Class

<GenerationTask("Show Model Sketch In View",
                "Show a model sketch in a specified drawing view",
                "embedded://DriveWorksPRGExtender.ShowSketch.png",
                "SOLIDWORKS PowerPack Drawing",
                GenerationTaskScope.Drawings,
                ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After Or ComponentTaskSequenceLocation.PreClose)>
Public Class ShowSketchInView
    Inherits GenerationTask

    Private Const VIEW_NAME As String = "ViewName"
    Private Const SKETCH_NAME As String = "SketchName"
    Private Const SHOW_OR_HIDE As String = "ShowOrHide"

    Private Const TASKNAME As String = "Show Model Sketch In View"

    Public Enum ShowOrHide
        SHOW = 1
        HIDE = 0
    End Enum
    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(VIEW_NAME,
                                                "Drawing View Name",
                                                "Name of the drawing view in which to show or hide the sketch",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(SKETCH_NAME,
                                                "Model Sketch Name",
                                                "Name of the model sketch to show or hide",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(SHOW_OR_HIDE,
                                                "Show/Hide",
                                                "SHOW to show the sketch or HIDE to hide the sketch",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(ShowOrHide), ShowOrHide.SHOW))
                }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        ' Validate inputs
        Dim sketchNameString As String = String.Empty
        Dim sketchNameList As New List(Of String)
        Dim viewNameString As String = String.Empty
        Dim viewNameList As New List(Of String)
        Dim showHide As ShowOrHide
        Dim showHideString As String = String.Empty
        If Not Me.Data.TryGetParameterValue(SKETCH_NAME, True, sketchNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, "Sketch name not provided")
            Return
        Else
            ' Create a list of sketch names
            sketchNameList = sketchNameString.Split("|").ToList()
        End If
        If Not Me.Data.TryGetParameterValue(VIEW_NAME, True, viewNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, "View name not provided")
            Return
        Else
            ' Create a list of view names
            viewNameList = viewNameString.Split("|").ToList()
        End If
        If Not Me.Data.TryGetParameterValue(Of ShowOrHide)(SHOW_OR_HIDE, showHide) Then
            ' Try to retrieve the value as a string
            If Not Me.Data.TryGetParameterValue(SHOW_OR_HIDE, True, showHideString) Then
                Me.SetExecutionResult(TaskExecutionResult.Failed, "Show or hide value not provided")
                Return
            Else
                ' Try to convert the string to a usable value
                Select Case showHideString.ToUpper()
                    Case "SHOW", "S", "TRUE"
                        showHide = ShowOrHide.SHOW
                    Case "HIDE", "H", "FALSE"
                        showHide = ShowOrHide.HIDE
                    Case Else
                        Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Show or hide value <{0}> was not supported. Use 'SHOW' or 'HIDE'.", showHideString))
                        Return
                End Select
            End If
        End If

        Dim app As SldWorks = Nothing
        Dim view As View = Nothing
        Dim refModel As ModelDoc2 = Nothing
        Dim drawComp As DrawingComponent = Nothing
        Dim selMgr As SelectionMgr = Nothing
        Dim sketchCount As Integer = 0
        Try
            ' Cycle through all of the views until we find the view that we're looking for
            view = model.Drawing.GetFirstView()
            ' The first view is the drawing sheet
            view = view.GetNextView
            While Not view Is Nothing
                If viewNameList.Contains(view.Name) Then
                    ' Get the referenced model and configuration for the view
                    refModel = view.ReferencedDocument
                    Dim configName As String = view.ReferencedConfiguration
                    ' Activate the reference document so that we can traverse its structure
                    Dim errorCode As Long
                    refModel = model.Application.Instance.ActivateDoc2(refModel.GetPathName, True, errorCode)
                    If Not refModel.ShowConfiguration2(configName) Then
                        Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Warning, "Model Generation Task", TASKNAME, "Unable to activate configuration", String.Format("Component: {0}; Configuration: {1}", refModel.GetPathName, configName))
                    End If
                    drawComp = view.RootDrawingComponent
                    'ToDo: Check for refModel and drawComp
                    ' Cycle through the features
                    Dim feat As Feature = refModel.FirstFeature
                    While Not feat Is Nothing
                        ' Check to see if this is a sketch
                        If feat.GetTypeName2() = "ProfileFeature" Then
                            ' Get the full sketch name
                            Dim fullSketchName = GetFullSketchName(view.Name, drawComp.Name, feat.Name)
                            ' Try to select the sketch
                            If model.Model.Extension.SelectByID2(fullSketchName, "SKETCH", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault) Then
                                If showHide = ShowOrHide.HIDE Then
                                    model.Model.BlankSketch()
                                    Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "Sketch hidden", String.Format("Sketch: {0}", fullSketchName))
                                Else
                                    model.Model.UnblankSketch()
                                    Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Model Generation Task", TASKNAME, "Sketch shown", String.Format("Sketch: {0}", fullSketchName))
                                End If
                            Else
                                warnings.Add(String.Format("Unable to select sketch <{0}> in view <{1}>", fullSketchName, view.Name))
                            End If
                        End If
                        feat = feat.GetNextFeature
                    End While
                    If drawComp.GetChildrenCount > 0 Then TraverseComp(drawComp, model.Model, view)
                    app.CloseDoc(refModel.GetPathName)
                End If
            End While
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If view IsNot Nothing Then view = Nothing
            If refModel IsNot Nothing Then refModel = Nothing
            If selMgr IsNot Nothing Then selMgr = Nothing
            If app IsNot Nothing Then app = Nothing
        End Try
        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub

    Private Function GetFullSketchName(ByVal viewName As String, ByVal viewCompName As String, ByVal featName As String) As String
        Dim comp As List(Of String) = viewCompName.Split("/").ToList()
        Dim sketchName As String = featName & "@" & comp(0) & "@" & viewName
        Dim path As String = String.Empty
        For index = 0 To comp.Count - 2
            sketchName = sketchName & "/" & comp(index + 1) & "@" & comp(index).Split("-")(0)
        Next
        Return sketchName
    End Function

    Private Sub TraverseComp(ByVal drawComp As DrawingComponent, modelDoc As ModelDoc2, view As View)

        Dim children As Object
        Dim tempComp As DrawingComponent

        If drawComp.GetChildrenCount > 0 Then
            children = drawComp.GetChildren
            For index = 0 To drawComp.GetChildrenCount - 1
                tempComp = children(index)
                TraverseComp(tempComp, modelDoc, view)
            Next
        End If
    End Sub

End Class

<GenerationTask("Select Model Sketch In View",
                "Select a model sketch in a specified drawing view",
                "embedded://DriveWorksPRGExtender.ShowSketch.png",
                "SOLIDWORKS PowerPack Drawing",
                GenerationTaskScope.Drawings,
                ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After Or ComponentTaskSequenceLocation.PreClose)>
Public Class SelectSketchInView
    Inherits GenerationTask

    Private Const VIEW_NAME As String = "ViewName"
    Private Const SKETCH_NAME As String = "SketchName"
    Private Const SELECT_ENTITIES As String = "SelectEntities"
    Private Const INCLUDE_CONST As String = "IncludeConstruction"

    Private Const TASKNAME As String = "Select Model Sketch In View"

    Private Enum GeomType
        FEATURE = 0
        SEGMENTS = 1
        GEOMETRY = 2
        CONSTRUCTION = 3
    End Enum

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(VIEW_NAME,
                                                "Drawing View Name",
                                                "Name of the drawing view in which to show or hide the sketch",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(SKETCH_NAME,
                                                "Model Sketch Name",
                                                "Name of the model sketch to show or hide",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(SELECT_ENTITIES,
                                                "Selection Type",
                                                "'FEATURE' to select the sketch feature (for exmaple, to show or hide the sketch), 'GEOMETRY' to select non-costruction sketch segments, 'CONSTRUCTION' to select construction sketch segments, or 'SEGMENTS' to select all sketch segments in the sketch (for example, to use Convert Entities), 'False' ",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual, GetType(GeomType), GeomType.FEATURE))
            }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        ' Validate inputs
        Dim sketchNameString As String = String.Empty
        Dim sketchNameList As New List(Of String)
        Dim viewNameString As String = String.Empty
        Dim viewNameList As New List(Of String)
        Dim selectEntities As GeomType
        If Not Me.Data.TryGetParameterValue(SKETCH_NAME, True, sketchNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, "Sketch name not provided")
            Return
        Else
            ' Create a list of sketch names
            sketchNameList = sketchNameString.Split("|").ToList()
        End If
        If Not Me.Data.TryGetParameterValue(VIEW_NAME, True, viewNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, "View name not provided")
            Return
        Else
            ' Create a list of view names
            viewNameList = viewNameString.Split("|").ToList()
        End If
        If Not Me.Data.TryGetParameterValue(Of GeomType)(SELECT_ENTITIES, selectEntities) Then
            warnings.Add("Unable to read Selection Type input. [Defaulting to FEATURE]")
            selectEntities = GeomType.FEATURE
        End If

        Dim app As SldWorks = Nothing
        Dim view As View = Nothing
        Dim refModel As ModelDoc2 = Nothing
        Dim drawComp As DrawingComponent = Nothing
        Dim selMgr As SelectionMgr = Nothing
        Dim sketchCount As Integer = 0
        Try
            ' Cycle through all of the views until we find the view that we're looking for
            view = model.Drawing.GetFirstView()
            ' The first view is the drawing sheet
            view = view.GetNextView
            While Not view Is Nothing
                If viewNameList.Contains(view.Name) Then
                    ' Get the referenced model and configuration for the view
                    refModel = view.ReferencedDocument
                    Dim configName As String = view.ReferencedConfiguration
                    ' Activate the reference document so that we can traverse its structure
                    Dim errorCode As Long
                    'refModel = model.Application.Instance.ActivateDoc2(refModel.GetPathName, True, errorCode)
                    If Not refModel.ShowConfiguration2(configName) Then
                        Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Warning, "Model Generation Task", TASKNAME, "Unable to activate configuration", String.Format("Component: {0}; Configuration: {1}", refModel.GetPathName, configName))
                    End If
                    drawComp = view.RootDrawingComponent
                    'ToDo: Check for refModel and drawComp
                    ' Cycle through the features
                    Dim feat As Feature = refModel.FirstFeature
                    While Not feat Is Nothing
                        ' Check to see if this is a sketch
                        If feat.GetTypeName2() = "ProfileFeature" And sketchNameList.Contains(feat.Name) Then
                            ' Get the full sketch name
                            Dim fullSketchName = GetFullSketchName(view.Name, drawComp.Name, feat.Name)
                            If selectEntities = GeomType.FEATURE Then
                                ' Try to select the sketch
                                If model.Model.Extension.SelectByID2(fullSketchName, "SKETCH", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault) Then
                                    If selMgr Is Nothing Then selMgr = model.Model.SelectionManager
                                    If selMgr IsNot Nothing Then
                                        If sketchNameList.Count <= selMgr.GetSelectedObjectCount2(-1) Then
                                            feat = Nothing
                                            Exit While
                                        End If
                                    End If
                                Else
                                    warnings.Add(String.Format("Unable to select sketch <{0}> in view <{1}>", fullSketchName, view.Name))
                                End If
                            Else
                                ' Try to get the sketch and cycle through its entities
                                Dim sketch As Sketch = feat.GetSpecificFeature2()
                                If Not sketch Is Nothing Then
                                    Dim sketchSegs() As Object = sketch.GetSketchSegments
                                    For segNum = LBound(sketchSegs) To UBound(sketchSegs)
                                        Dim seg As SketchSegment = sketchSegs(segNum)
                                        If seg IsNot Nothing Then
                                            If selectEntities = GeomType.SEGMENTS _
                                                Or (selectEntities = GeomType.CONSTRUCTION And seg.ConstructionGeometry) _
                                                Or (selectEntities = GeomType.GEOMETRY And Not (seg.ConstructionGeometry)) Then
                                                If model.Model.Extension.SelectByID2(seg.GetName() & "@" & fullSketchName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault) Then
                                                    sketchCount += 1
                                                Else
                                                    warnings.Add(String.Format("Unable to select sketch segment {0}", seg.GetName()))
                                                End If
                                            End If
                                        End If
                                    Next segNum
                                End If
                            End If
                        End If
                        feat = feat.GetNextFeature
                    End While
                    If drawComp.GetChildrenCount > 0 Then TraverseComp(drawComp, model.Model, view)
                    model.Application.Instance.CloseDoc(refModel.GetPathName)
                    ' Close objects that we opened
                    view = Nothing
                    drawComp = Nothing
                    refModel = Nothing
                    selMgr = Nothing
                    Exit While
                End If
            End While
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If view IsNot Nothing Then view = Nothing
            If refModel IsNot Nothing Then refModel = Nothing
            If selMgr IsNot Nothing Then selMgr = Nothing
            If app IsNot Nothing Then app = Nothing
        End Try
        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub

    Private Function GetFullSketchName(ByVal viewName As String, ByVal viewCompName As String, ByVal featName As String) As String
        Dim comp As List(Of String) = viewCompName.Split("/").ToList()
        Dim sketchName As String = featName & "@" & comp(0) & "@" & viewName
        Dim path As String = String.Empty
        For index = 0 To comp.Count - 2
            sketchName = sketchName & "/" & comp(index + 1) & "@" & comp(index).Split("-")(0)
        Next
        Return sketchName
    End Function

    Private Sub TraverseComp(ByVal drawComp As DrawingComponent, modelDoc As ModelDoc2, view As View)

        Dim children As Object
        Dim tempComp As DrawingComponent

        If drawComp.GetChildrenCount > 0 Then
            children = drawComp.GetChildren
            For index = 0 To drawComp.GetChildrenCount - 1
                tempComp = children(index)
                TraverseComp(tempComp, modelDoc, view)
            Next
        End If
    End Sub

End Class

<GenerationTask("Rename Feature in Part",
                "Change the name of an existing feature in a part model",
                "embedded://DriveWorksPRGExtender.ShowSketch.png",
                "PRG Tools",
                GenerationTaskScope.Parts,
                ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After Or ComponentTaskSequenceLocation.PreClose)>
Public Class RenameFeature
    Inherits GenerationTask

    Private Const FEATURE_NAME As String = "FeatureName"
    Private Const NEW_NAME As String = "NewName"

    Private Const TASKNAME As String = "Rename Feature"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(FEATURE_NAME,
                                                "Current Feature Name",
                                                "Name of the feature to select and rename",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(NEW_NAME,
                                                "New Feature Name",
                                                "New name of the feature",
                                                "Inputs",
                                                New ComponentTaskParameterMetaData(PropertyBehavior.StandardOptionDynamicManual))
            }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        Dim warnings As New List(Of String)
        ' Validate inputs
        Dim featName As String = String.Empty
        Dim sketchNameList As New List(Of String)
        Dim newName As String = String.Empty
        Dim viewNameList As New List(Of String)
        If Not Me.Data.TryGetParameterValue(FEATURE_NAME, True, featName) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, "Current feature name not provided")
            Return
        ElseIf Not Me.Data.TryGetParameterValue(NEW_NAME, True, newName) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, "New name not provided")
            Return
        End If

        Dim app As SldWorks = Nothing
        Dim sketchCount As Integer = 0
        Try
            Dim feat As Feature = model.Model.FirstFeature
            ' Should we check through all of the features to get the new name?
            While Not feat Is Nothing
                ' Check to see if this is the feature
                If featName.Equals(feat.Name, StringComparison.CurrentCultureIgnoreCase) Then
                    feat.Select(False)
                    Dim featMgr As FeatureManager = model.Model.FeatureManager
                    Dim featData As Object = feat.GetSpecificFeature()
                    feat.Name = newName
                    feat.ModifyDefinition(featData, model.Model, Nothing)
                    If Not (model.Model.EditRebuild3()) Then
                        warnings.Add("Rebuild Error after renaming feature")
                    End If
                    Exit While
                End If
                feat = feat.GetNextFeature
            End While
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception Error: {0}", ex.Message))
        Finally
            ' Clean Up
            If app IsNot Nothing Then app = Nothing
        End Try
        If warnings.Count = 0 Then
            Me.SetExecutionResult(TaskExecutionResult.Success, String.Format("{0} executed with no warnings", TASKNAME))
        Else
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Warnings: " & String.Join(";", warnings.ToArray()))
        End If
    End Sub
End Class
'ToDo: Create "SELECT ALL ENTITIES ON LAYER" - Hide all other layers, then do a select all visible
'ToDo: Create "CONVERT SELECTION TO DRAWING" - Inputs: None - Do ConvertEntities on the sketch 
'ToDo: Create "SELECT ENTITIES BY NAME" - Input: ListOfItems, ViewName
'ToDo: Create "HAS ITEMS SELECTED" - True if the SelectedItemCount > 0