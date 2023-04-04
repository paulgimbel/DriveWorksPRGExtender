Imports DriveWorks.Components
Imports DriveWorks.Components.Tasks
Imports DriveWorks.Forms
Imports DriveWorks.Reporting
Imports DriveWorks.SolidWorks
Imports DriveWorks.SolidWorks.Generation
Imports DriveWorks.SolidWorks.Generation.Proxies
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

<GenerationTask("Insert A Revision Table",
           "Add a Revision to the specified sheer of the drawing",
           "embedded://DriveWorksPRGExtender.Triangle.bmp",
           "PRG Tasks", GenerationTaskScope.Drawings, ComponentTaskSequenceLocation.After)>
Public Class InsertRevisionTable
    Inherits GenerationTask

    Private Const SHEET_NAME As String = "SheetName"
    Private Const ANCHOR_POINT As String = "AnchorPoint"
    Private Const TEMPLATE_PATH As String = "TemplatePath"
    Private Const SYMBOL_SHAPE As String = "SymbolShape"

    Private TASKNAME As String = "Insert Revision Table"

    Private Enum AnchorPoint
        TopRight = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight
        TopLeft = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
        BottomRight = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight
        BottomLeft = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomLeft
    End Enum

    Private Enum RevSymbolShape
        Circle = swRevisionTableSymbolShape_e.swRevisionTable_CircleSymbol
        Hexagon = swRevisionTableSymbolShape_e.swRevisionTable_HexagonSymbol
        Square = swRevisionTableSymbolShape_e.swRevisionTable_SquareSymbol
        Triangle = swRevisionTableSymbolShape_e.swRevisionTable_TriangleSymbol
    End Enum

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(SHEET_NAME,
                                               "Sheet Name",
                                               "Pipe-delimited list of drawing sheet names to receive revision tables (use * to create tables on all sheets)",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(DataModel.PropertyBehavior.StandardOptionDynamicManual)),
                New ComponentTaskParameterInfo(ANCHOR_POINT,
                                               "Anchor Point",
                                               "Drawing sheet anchor point to use for locating the revision table",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(DataModel.PropertyBehavior.StandardOptionDynamicManual,
                                                                                  GetType(AnchorPoint), AnchorPoint.TopRight)),
                New ComponentTaskParameterInfo(SYMBOL_SHAPE,
                                               "Sheet Name",
                                               "Shape of the revision symbol to be create by the revision table",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(DataModel.PropertyBehavior.StandardOptionDynamicManual,
                                                                                  GetType(RevSymbolShape), RevSymbolShape.Triangle)),
                New ComponentTaskParameterInfo(TEMPLATE_PATH,
                                               "Sheet Name",
                                               "Pipe-delimited list of drawing sheet names to receive revision tables (use * to create tables on all sheets)",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(DataModel.PropertyBehavior.StandardOptionDynamicManual))
            }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        ' Validate Inputs
        Dim sheetNameString As String = String.Empty
        Dim sheetNameList As New List(Of String)
        Dim templateName As String = String.Empty
        Dim anchorPoint As AnchorPoint
        Dim revSymbolShape As RevSymbolShape
        ' Verify that we got data in all of the inputs
        If Not Me.Data.TryGetParameterValue(SHEET_NAME, True, sheetNameString) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No value was provided for the drawing sheet name")
            Return
        ElseIf Not Me.Data.TryGetParameterValue(TEMPLATE_PATH, True, templateName) Then
            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "No value was provided for the template path")
            Return
        End If
        If Not Me.Data.TryGetParameterValue(Of AnchorPoint)(ANCHOR_POINT, anchorPoint) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Generation Task", TASKNAME, "No anchor point provided", "Default value will be used")
            anchorPoint = AnchorPoint.TopRight
        End If
        If Not Me.Data.TryGetParameterValue(Of RevSymbolShape)(SYMBOL_SHAPE, revSymbolShape) Then
            Me.Report.WriteEntry(ReportingLevel.Verbose, ReportEntryType.Information, "Generation Task", TASKNAME, "No revision symbol shape provided", "Default value will be used")
            Return
        End If

        Dim swApp As SldWorks = Nothing
        Dim currentSheet As Sheet = Nothing
        Dim revTblAnn As RevisionTableAnnotation = Nothing
        Dim draw As DrawingDoc = Nothing
        Try
            draw = model.Drawing
            For Each sheetName In sheetNameList
                ' Switch to the named sheet
                currentSheet = draw.GetCurrentSheet
                draw.ActivateSheet(sheetName)
                revTblAnn = currentSheet.InsertRevisionTable2(True, 0, 0, swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, templateName, swRevisionTableSymbolShape_e.swRevisionTable_CircleSymbol, True)
            Next
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception: {0}", ex.Message))
        Finally
            If swApp IsNot Nothing Then swApp = Nothing
            If currentSheet IsNot Nothing Then currentSheet = Nothing
            If revTblAnn IsNot Nothing Then revTblAnn = Nothing
            If draw IsNot Nothing Then draw = Nothing
        End Try
    End Sub
End Class

<GenerationTask("Delete All Rows From Revision Table",
           "Delete all data rows from a Revision table on the drawing",
           "embedded://DriveWorksPRGExtender.Triangle.bmp",
           "SOLIDWORKS PowerPack Drawing",
           GenerationTaskScope.Drawings,
           ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After)>
Public Class ClearRevisionTable
    Inherits GenerationTask

    Private Const SHEET_NAME As String = "SheetName"

    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            Return New ComponentTaskParameterInfo() {
                New ComponentTaskParameterInfo(SHEET_NAME,
                                               "Sheet Name",
                                               "Pipe-delimited list of drawing sheet names with the revision table to clear (use * to clear tables on all sheets)",
                                               "Inputs",
                                               New ComponentTaskParameterMetaData(DataModel.PropertyBehavior.StandardOptionDynamicManual))
                }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)
        ' Validate Inputs
        Dim sheetNameString As String = String.Empty
        Dim sheetNameList As New List(Of String)

        If Not Me.Data.TryGetParameterValue(SHEET_NAME, sheetNameString) Then
            ' No sheet name was received, can't go on
            Me.SetExecutionResult(TaskExecutionResult.Failed, "No sheet name was found")
            Return
        Else
            sheetNameList = sheetNameString.Split("|").ToList()
        End If

        Dim swApp As SldWorks = Nothing
        Dim currentSheet As Sheet = Nothing
        Dim revTblFeat As RevisionTableFeature = Nothing
        Dim revTblAnn As RevisionTableAnnotation = Nothing
        Dim tableAnn As TableAnnotation = Nothing
        Dim draw As DrawingDoc = Nothing
        Try
            swApp = model.Application.Instance
            draw = model.Drawing
            If draw Is Nothing Then
                Me.SetExecutionResult(TaskExecutionResult.Failed, "Unable to retrieve the drawing")
                Return
            End If
            If sheetNameString.Contains("*") Then sheetNameList = draw.GetSheetNames.ToList()
            For Each sheetName In sheetNameList
                currentSheet = draw.Sheet(sheetName)
                If currentSheet Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, String.Format("Unable to retrieve sheet {0}", sheetName))
                    Return
                End If
                revTblAnn = currentSheet.RevisionTable
                If revTblAnn Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to retrieve a revision table on the current sheet")
                    Return
                End If
                tableAnn = revTblAnn
                If tableAnn Is Nothing Then
                    Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, "Unable to retrieve a table annotation for the rev table")
                    Return
                End If
                For rowNum = tableAnn.RowCount - 1 To 0 Step -1
                    Dim rowID As Integer = revTblAnn.GetIdForRowNumber(rowNum)
                    If Not (rowID = 0) Then
                        If Not revTblAnn.DeleteRevision(rowID, True) Then
                            Me.SetExecutionResult(TaskExecutionResult.SuccessWithWarnings, String.Format("Unable to delete revision ID {0}", rowID))
                        End If
                    End If
                Next
            Next
            Me.SetExecutionResult(TaskExecutionResult.Success)
        Catch ex As Exception
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Exception: {0}", ex.Message))
        Finally
            If swApp IsNot Nothing Then swApp = Nothing
            If currentSheet IsNot Nothing Then currentSheet = Nothing
            If revTblFeat IsNot Nothing Then revTblFeat = Nothing
            If revTblAnn IsNot Nothing Then revTblAnn = Nothing
            If tableAnn IsNot Nothing Then tableAnn = Nothing
            If draw IsNot Nothing Then draw = Nothing
        End Try


    End Sub

End Class