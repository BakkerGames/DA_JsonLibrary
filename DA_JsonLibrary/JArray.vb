' --- Purpose: Provide a JSON Array class
' --- Author : Scott Bakker
' --- Created: 09/13/2019
' --- LastMod: 04/06/2020

' --- Notes  : The values in the list ARE ordered based on when they are added.
' ---        : The values are NOT sorted, and there can be duplicates.
' ---        : The function ToStringFormatted() will return a string representation with
' ---          whitespace added. Two spaces are used for indenting, and CRLF between lines.

Imports System.Text
Imports DA_JsonLibrary.JsonRoutines

Public Class JArray

    Implements IEnumerable(Of Object)

    Private _data As List(Of Object)

    Public Sub New()
        ' --- Purpose: Initialize the internal structures of a new JArray
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        _data = New List(Of Object)
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of Object) Implements IEnumerable(Of Object).GetEnumerator
        ' --- Purpose: Provide IEnumerable access directly to _data
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return DirectCast(Me._data, IEnumerable(Of Object)).GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        ' --- Purpose: Provide IEnumerable access directly to _data
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return DirectCast(Me._data, IEnumerable(Of Object)).GetEnumerator()
    End Function

    Public Sub Add(ByVal value As Object)
        ' --- Purpose: Adds a new value to the end of the JArray list
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        ' --- Changes: 10/03/2019 Removed extra string processing, was wrong
        _data.Add(value)
    End Sub

    Public Sub Append(ByVal list As IEnumerable)
        ' --- Purpose: Append all values in the sent IEnumerable at the end of the JArray list
        ' --- Author : Scott Bakker
        ' --- Created: 02/12/2020
        If list IsNot Nothing Then
            For Each obj As Object In list
                _data.Add(obj)
            Next
        End If
    End Sub

    Public Function Count() As Integer
        ' --- Purpose: Return the count of items in the JArray
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return _data.Count
    End Function

    Default Public Property Item(ByVal index As Integer) As Object
        ' --- Purpose: Give access to item values by index
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Get
            If index < 0 OrElse index >= _data.Count Then
                Throw New IndexOutOfRangeException
            End If
            Return _data(index)
        End Get
        Set(value As Object)
            If index < 0 OrElse index >= _data.Count Then
                Throw New IndexOutOfRangeException
            End If
            _data(index) = value
        End Set
    End Property

    Public Function Items() As List(Of Object)
        ' --- Purpose: Get cloned list of all objects
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return New List(Of Object)(_data)
    End Function

    Public Sub RemoveAt(ByVal index As Integer)
        ' --- Purpose: Remove the item at the specified index
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If index < 0 OrElse index >= _data.Count Then
            Throw New IndexOutOfRangeException
        End If
        _data.RemoveAt(index)
    End Sub

#Region "ToString"

    Public Overrides Function ToString() As String
        ' --- Purpose: Convert this JArray into a string
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        ' --- Notes  : This could be implemented as ToStringFormatted(-1) but
        ' ---          it is separate to get better performance.
        Dim result As New StringBuilder
        result.Append("[")
        Dim addComma As Boolean = False
        For Each value As Object In _data
            If addComma Then
                result.Append(",")
            Else
                addComma = True
            End If
            result.Append(ValueToString(value))
        Next
        result.Append("]")
        Return result.ToString()
    End Function

    Public Function ToStringFormatted() As String
        ' --- Purpose: Convert this JArray into a string with formatting
        ' --- Author : Scott Bakker
        ' --- Created: 10/17/2019
        Dim indentlevel As Integer = 0
        Return ToStringFormatted(indentlevel)
    End Function

    Friend Function ToStringFormatted(ByRef indentLevel As Integer) As String
        ' --- Purpose: Convert this JArray into a string with formatting
        ' --- Author : Scott Bakker
        ' --- Created: 10/17/2019
        If _data.Count = 0 Then
            Return "[]" ' avoid indent errors
        End If
        Dim result As New StringBuilder
        result.Append("[")
        If indentLevel >= 0 Then
            result.AppendLine()
            indentLevel += 1
        End If
        Dim addComma As Boolean = False
        For Each value As Object In _data
            If addComma Then
                result.Append(",")
                If indentLevel >= 0 Then
                    result.AppendLine()
                End If
            Else
                addComma = True
            End If
            If indentLevel > 0 Then
                result.Append(IndentSpace(indentLevel))
            End If
            result.Append(ValueToString(value, indentLevel))
        Next
        If indentLevel >= 0 Then
            result.AppendLine()
            If indentLevel > 0 Then
                indentLevel -= 1
            End If
            result.Append(IndentSpace(indentLevel))
        End If
        result.Append("]")
        Return result.ToString()
    End Function

#End Region

#Region "Parse"

    Public Shared Function Parse(ByVal value As String) As JArray
        ' --- Purpose: Convert a string into a JArray
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Dim pos As Integer = 0
        Return Parse(value, pos)
    End Function

    Friend Shared Function Parse(ByVal value As String, ByRef pos As Integer) As JArray
        ' --- Purpose: Convert a partial string into a JArray
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If value Is Nothing OrElse value.Length = 0 Then
            Return Nothing
        End If
        Dim result As New JArray
        Dim tempValue As String
        SkipWhitespace(value, pos)
        If value(pos) <> "[" Then
            Throw New SystemException($"JSON Error: Unexpected token to start JArray: {value(pos)}")
        End If
        pos += 1
        Do
            SkipWhitespace(value, pos)
            ' --- check for symbols
            If value(pos) = "]" Then
                pos += 1
                Exit Do ' --- done building JArray
            End If
            If value(pos) = "," Then
                ' --- this logic ignores extra commas, but is ok
                pos += 1
                Continue Do ' --- Next value
            End If
            ' --- Check for JArray, JArray
            If value(pos) = "{" Then
                Dim jo As JObject = JObject.Parse(value, pos)
                result.Add(jo)
                Continue Do
            End If
            If value(pos) = "[" Then
                Dim ja As JArray = Parse(value, pos)
                result.Add(ja)
                Continue Do
            End If
            ' --- Get value as a string, convert to object
            tempValue = GetToken(value, pos)
            result.Add(JsonValueToObject(tempValue))
        Loop
        Return result
    End Function

#End Region

#Region "Clone"

    Public Shared Function Clone(ByVal ja As JArray) As JArray
        ' --- Purpose: Clones a JArray
        ' --- Author : Scott Bakker
        ' --- Created: 09/20/2019
        Dim result As New JArray
        If ja IsNot Nothing AndAlso ja._data IsNot Nothing Then
            result._data = New List(Of Object)(ja._data)
        End If
        Return result
    End Function

#End Region

End Class
