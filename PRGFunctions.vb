Imports DriveWorks
Imports DriveWorks.Applications
Imports DriveWorks.Extensibility
Imports DriveWorks.Forms
Imports DriveWorks.Navigation
Imports SolidWorks.Interop.sldworks
Imports Titan.Rules.Execution
Public Class PRGFunctions
    Inherits SharedProjectExtender


    '' <summary>
    '' Takes the name of a form or list of forms as an input and returns a boolean determining whether any controls on that form or those forms are in error
    '' </summary>
    '' <param name="formName">Form name or pipe-delimited list of form names</param>
    '' <returns>Boolean - True if any control error rules are non-blank and non-zero</returns>

    <Udf()>
    <FunctionInfo("Identify if a form has any controls in error", "PRG")>
    Function PRGFormHasError(<ParamInfo("Form Name", "Name Of ")> ByVal formName As String) As Boolean

        Dim ErrorFound As Boolean = False

        ' Split the formName in case it's a list
        Dim formsList As List(Of String) = formName.Split("|").ToList()
        ' Try to get the form name
        Dim targetForm As FormNavigationStep
        Dim forms() As FormNavigationStep = Me.Project.Navigation.GetForms(True, True)
        For Each form As FormNavigationStep In forms
            If formsList.Contains(form.Name) Then
                targetForm = form
                ' Cycle through the controls on that form
                For Each control As ControlBase In targetForm.Form.GetAllControls()
                    'If has an error, change result
                    If control.MessageCode = "0" Or String.IsNullOrEmpty(control.MessageCode) Then
                        ' No error here
                    Else
                        ErrorFound = True
                        ' No sense looping any further, we found one.
                        Return True
                    End If
                Next control
            End If
        Next form

        Return ErrorFound

    End Function


    ''' <summary>
    ''' Takes the name of a form in as an input and returns a count of the number of controls on that form or forms are in error
    ''' </summary>
    ''' <param name="formName">Form name or pipe-delimited list of form names</param>
    ''' <returns>Integer representing the number of controls with error rules are non-blank and non-zero</returns>

    <Udf()>
    <FunctionInfo("Count the number of errors on a form or forms", "PRG")>
    Function PRGFormCountErrors(<ParamInfo("Form Name", "Name Of ")> ByVal formName As String) As Double

        Dim ErrorsFound As Double = 0

        ' Split the formName in case it's a list
        Dim formsList As List(Of String) = formName.Split("|").ToList()
        ' Try to get the form name
        Dim targetForm As FormNavigationStep = Nothing
        Dim forms() As FormNavigationStep = Me.Project.Navigation.GetForms(True, True)
        For Each form As FormNavigationStep In forms
            If formsList.Contains(form.Name) Then
                targetForm = form
                ' Cycle through the controls on that form
                For Each control As ControlBase In targetForm.Form.GetAllControls()
                    'If has an error, change result
                    If control.MessageCode = "0" Or String.IsNullOrEmpty(control.MessageCode) Then
                        ' No error here
                    Else
                        ErrorsFound = +1
                    End If
                Next control
                Exit For
            End If
        Next form

        Return ErrorsFound

    End Function

    <Udf()>
    <FunctionInfo("Count the number of times a string occurs in a list", "PRG")>
    Function PRGListCountInstances(<ParamInfo("List", "String containing the pipe-delimited list")> ByVal InputString As String, <ParamInfo("Search String", "String to count in the list")> ByVal SearchString As String) As Double
        Dim result As Integer

        Dim TempArray() As String

        TempArray = Split(InputString, "|")

        result = 0

        For i = 0 To UBound(TempArray)
            If TempArray(i) = SearchString Then
                Dim instanceCount As Integer = result + 1
                result = instanceCount
            End If
        Next i

        'Return the result
        Return result
    End Function

    <Udf()>
    <FunctionInfo("Conditional function to check a list for an item, returning TRUE or FALSE (not case sensitive)", "PRG")>
    Public Function PRGListContains(<ParamInfo("Search string", "Item to find in the list")> ByVal searchString As String,
                               <ParamInfo("List", "Pipe-delimited list of values")> ByVal listString As String) As Boolean

        Dim result As Boolean = False
        Dim listAsCollection As List(Of String) = listString.Split("|").ToList
        result = listAsCollection.Contains(searchString)

        Return result

    End Function

    '<Udf()>
    '<FunctionInfo("Returns the name of the current DriveWorks Group")>
    'Public Function GetGroupName() As String

    '    Return Me.Project.Group.Name

    'End Function

    <Udf()>
    <FunctionInfo("Returns a list of the names of all controls in the current project", "PRG")>
    Public Function PRGGetControlNames(<ParamInfo("Include Non-Inputs", "Include non-input controls (ex. labels, picutre boxes)")> ByVal IncludeNonInputs As Boolean) As String

        Dim controlNames As New List(Of String)

        Dim forms() As FormNavigationStep = Me.Project.Navigation.GetForms(True, True)
        For Each formStep As FormNavigationStep In forms
            Dim controls() As ControlBase = formStep.Form.GetAllControls
            For Each control As ControlBase In controls
                If IncludeNonInputs Or control.HasInputValue Then
                    controlNames.Add(control.Name)
                End If
            Next control
        Next formStep

        Return String.Join("|", controlNames.ToArray)

    End Function

    <Udf()>
    <FunctionInfo("Turns a DriveWorks Array into HTML table code", "PRG")>
    Public Function PRGTableToHTML(<ParamInfo("Table", "DriveWorks array")> ByVal InputTable As IArrayValue) As String

        Dim output As String = String.Empty

        Dim TableData(,) As Object = InputTable.ToArray

        output = String.Format("<table>")
        For i = TableData.GetLowerBound(0) To TableData.GetUpperBound(0)
            output = String.Format("{0}<tr>", output)
            For j = TableData.GetLowerBound(1) To TableData.GetUpperBound(1)
                output = String.Format("{0}<td>{1}</td>", output, TableData(i, j).ToString())
            Next j
            output = String.Format("{0}</tr>", output)
        Next i
        output = String.Format("{0}</table>", output)

        Return output
    End Function

    <Udf()>
    <FunctionInfo("Returns values from a table row as a pipe-delimited list", "PRG")>
    Public Function PRGTableRowToList(<ParamInfo("Table", "DriveWorks array")> ByVal InputTable As IArrayValue, <ParamInfo("Row number", "Number of the row to return, use 0 for the header row")> RowNumber As Double) As String

        Dim valueList As New List(Of String)
        Dim listString As String = String.Empty

        ' Extract the data from the table
        Dim TableData(,) As Object = InputTable.ToArray

        ' Check to make sure that the number is a valid row
        If RowNumber >= TableData.GetLowerBound(0) And RowNumber <= TableData.GetUpperBound(0) Then
        Else
            Return listString
        End If

        For i = TableData.GetLowerBound(1) To TableData.GetUpperBound(1)
            valueList.Add(TableData(RowNumber, i))
        Next i

        ' Convert the list to a pipe-delimited string
        listString = String.Join("|", valueList.ToArray)
        Return listString
    End Function

    <Udf()>
    <FunctionInfo("Combines contiguous values in a list", "PRG")>
    Public Function PRGListCombineValues(<ParamInfo("List", "Pipe-Delimited list of values")> ByVal InputList As String, <ParamInfo("Total Values", "When true, will try to sum numeric values")> ByVal Total As Boolean) As String

        Dim returnList As String = String.Empty

        ' Make sure that they didn't pass us an empty string
        If String.IsNullOrEmpty(InputList) Then Return String.Empty

        Dim values As List(Of String) = InputList.Split("|").ToList
        ' Not sure if we need to check to see if the list converted before we try to access anything, but I'm going to do it anyway just in case we get paid by the line of code
        If values.Count = 0 Then Return String.Empty
        If values.Count = 1 Then Return values(0)

        Dim NewValues As New List(Of String)

        Dim cachedValue As String = values(0)
        For i = 1 To values.Count - 1
            ' Check to see if the next value is blank. If it is, then clear out the cached value, if there is one, and write the total to the new list. If it's not, try to add it to the cache
            If String.IsNullOrEmpty(values(i)) Or values(i) = "0" Then
                ' Make sure that we're not going to write an empty string to the new array
                If Not (String.IsNullOrEmpty(cachedValue)) And Not (cachedValue = "0") Then
                    NewValues.Add(cachedValue)
                    cachedValue = String.Empty
                End If
            Else
                ' Check to see if the TOTAL switch is thrown to see if we need to try to add or just concatenate
                Dim cachedNumber, newNumber As Double
                If Total And Double.TryParse(cachedValue, cachedNumber) And Double.TryParse(values(i), newNumber) Then
                    cachedNumber += newNumber
                    cachedValue = cachedNumber.ToString
                Else
                    cachedValue += values(i)
                End If
            End If
        Next i
        ' If we have a value in the cached value at the end of the array, we need to add it to the new array
        If Not (String.IsNullOrEmpty(cachedValue)) Then NewValues.Add(cachedValue)

        returnList = String.Join("|", NewValues.ToArray)
        Return returnList
    End Function

    <Udf()>
    <FunctionInfo("Combines contiguous values in a list", "PRG")>
    Public Function PRGTableFromString(<ParamInfo("TableString", "Formatted string")> ByVal InputTable As String, <ParamInfo("Row Delimiter", "Character used to separate rows (semi-colon by default)")> Optional ByVal RowDelimiter As String = ";", <ParamInfo("Column Delimiter", "Character used to separate column values (comma by default)")> Optional ByVal ColumnDelimiter As String = ",") As IArrayValue

        Dim resultTable As IArrayValue = Nothing

        If InputTable.StartsWith("{") And InputTable.EndsWith("}") Then ' remove the characters
            InputTable = InputTable.Substring(1, InputTable.Length - 2)
        End If

        Dim inputRows As List(Of String) = InputTable.Split(RowDelimiter).ToList

        For Each row As String In inputRows
            Dim colValues As List(Of String) = row.Split(ColumnDelimiter).ToList
            For Each val As String In colValues
            Next val
        Next row

        Return resultTable

    End Function

    <Udf()>
    <FunctionInfo("Tries to make an IArray recognizable", "PRG")>
    Public Function PRGRecognizeIArray(<ParamInfo("Array", "Formatted array")> ByVal InputArray As String) As IArrayValue

        ' Strip off the surrounding braces
        If InputArray.StartsWith("{") And InputArray.EndsWith("}") Then InputArray = InputArray.Substring(1, InputArray.Length - 2)


        Dim StringArray As String = "{"

        ' Split the data into rows
        Dim rows() As String = InputArray.Split(";")
        Dim row() As String = rows(0).Split(",")
        Dim rowCount As Integer = rows.GetUpperBound(0)
        Dim colCount As Integer = row.GetUpperBound(0)
        Dim inputAsArray(rowCount, colCount) As String

        For i = rows.GetLowerBound(0) To rows.GetUpperBound(0)
            Dim vals() As String = rows(i).Split(",")
            For j = vals.GetLowerBound(0) To vals.GetUpperBound(0)
                StringArray = String.Format("{0}{1}{2}{3}{4}", StringArray, IIf(j = vals.GetLowerBound(0), "", ","), "", vals(j).ToString, "")  ' Removed the quotation marks around the value for {2} and {4}
                inputAsArray(i, j) = vals(j).Replace("""", "")
            Next j
            StringArray = String.Format("{0}{1}", StringArray, IIf(i = rows.GetUpperBound(0), "}", ";"))
        Next i

        Dim stdArray As New StandardArrayValue(inputAsArray)

        Return stdArray

    End Function


    '<Udf()>
    '<FunctionInfo("Multi-level Sort of a Table", "PRG")>
    'Public Function PRGTableSort(<ParamInfo("Table", "Unsorted Table")> ByVal inputTable As IArrayValue, <ParamInfo("Columns", "Column names or indices to sort")> ByVal Columns As String()) As String

    '    If inputTable Is Nothing Or inputTable.Rows <= 0 Or inputTable.Columns <= 0 Then
    '        Return String.Empty
    '    End If

    '    Dim comparison = StringComparison.Ordinal

    '    ' ToDo: Add in ignore case on the column names?



    '    ' Get Column Indices
    '    Dim colIndices As New List(Of Integer)
    '    For Each colInput As String In Columns
    '        Dim columnFound As Boolean = False
    '        For colIndex As Integer = 0 To inputTable.Columns - 1
    '            Dim colHeader = inputTable.GetElementAsString(Globalization.CultureInfo.CurrentUICulture, 0, colIndex)
    '            If String.Equals(colHeader, colInput, comparison) Then
    '                colIndices.Add(colIndex)
    '                columnFound = True
    '                Exit For
    '            Else

    '            End If
    '        Next
    '    Next

    '    ' Filter the table



    '    Dim sortedTable As IArrayValue

    '    Return sortedTable

    'End Function


    <Udf()>
    <FunctionInfo("Get start of a node name", "PRG")>
    Public Function PRGGetParent(<ParamInfo("Node", "Full name of the node")> ByVal nodeName As String, <ParamInfo("Levels", "Number of levels up to go")> ByVal levels As Double) As String

        Dim nodeList As List(Of String) = nodeName.Split("\").ToList
        Dim newNodeName As String = String.Empty
        For index = 1 To nodeList.Count - levels
            newNodeName = String.Format("{0}{1}{2}", newNodeName, IIf(String.IsNullOrEmpty(newNodeName), "", "\"), nodeList(index - 1))
        Next
        Return newNodeName
    End Function

    <Udf()>
    <FunctionInfo("Get end of a node name", "PRG")>
    Public Function PRGGetChild(<ParamInfo("Node", "Full name of the node")> ByVal nodeName As String, <ParamInfo("Levels", "Number of levels up to return")> ByVal levels As Double) As String

        Dim nodeList As List(Of String) = nodeName.Split("\").ToList
        Dim newNodeName As String = String.Empty
        For index = nodeList.Count - levels To nodeList.Count
            newNodeName = String.Format("{0}{1}{2}", newNodeName, IIf(String.IsNullOrEmpty(newNodeName), "", "\"), nodeList(index - 1))
        Next
        Return newNodeName
    End Function


End Class
