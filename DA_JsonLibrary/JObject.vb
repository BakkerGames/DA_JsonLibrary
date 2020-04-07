' --- Purpose: Provide a JSON Object class
' --- Author : Scott Bakker
' --- Created: 09/13/2019
' --- LastMod: 04/06/2020

' --- Notes  : The keys in this JObject implementation are case sensitive, so "abc" <> "ABC".
' ---        : Keys cannot be null, empty, or contain only whitespace.
' ---        : The items in this JObject are NOT ordered in any way. Specifically, successive
' ---          calls to ToString() may not return the same results.
' ---        : The function ToStringSorted() may be used to return a sorted list, but will be
' ---          somewhat slower due to overhead. The ordering is not specified here but it
' ---          should be consistent across calls.
' ---        : The function ToStringFormatted() will return a string representation with
' ---          whitespace added. Two spaces are used for indenting, and CRLF between lines.

Imports System.Text
Imports DA_JsonLibrary.JsonRoutines

Public Class JObject

    Implements IEnumerable(Of String)

    Private _data As Dictionary(Of String, Object)

    Public Sub New()
        ' --- Purpose: Initialize the internal structures of a new JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        _data = New Dictionary(Of String, Object)
    End Sub

    Public Sub New(jo As JObject)
        ' --- Purpose: Initialize the internal structures of a new JObject with values
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        _data = New Dictionary(Of String, Object)
        Merge(jo)
    End Sub

    Public Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        ' --- Purpose: Provide IEnumerable access directly to _data.Keys
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return DirectCast(_data.Keys, IEnumerable(Of String)).GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        ' --- Purpose: Provide IEnumerable access directly to _data.Keys
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return DirectCast(_data.Keys, IEnumerable(Of String)).GetEnumerator()
    End Function

    Public Sub Add(ByVal key As String, ByVal value As Object)
        ' --- Purpose: Adds a new key/value pair to JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        ' --- Changes: 10/03/2019 Removed extra string processing, was wrong
        ' --- Notes  : Throws an error if the key already exists. Use .Item(key) = value when that happens.
        If IsWhitespaceString(key) Then
            Throw New ArgumentNullException(NameOf(key), "JSON Error: Key cannot be null/empty/whitespace")
        End If
        If _data.ContainsKey(key) Then
            Throw New SystemException($"JSON Error: Key already exists: {key}")
        End If
        _data.Add(key, value)
    End Sub

    Public Sub Clear()
        ' --- Purpose: Clears all items from the current JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        _data.Clear()
    End Sub

    Public Function Contains(ByVal key As String) As Boolean
        ' --- Purpose: Identifies whether a key exists in the current JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If IsWhitespaceString(key) Then
            Throw New ArgumentNullException(NameOf(key), "JSON Error: Key cannot be null/empty/whitespace")
        End If
        Return _data.ContainsKey(key)
    End Function

    Public Function Count() As Integer
        ' --- Purpose: Return the count of items in the JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Return _data.Count
    End Function

    Default Public Property Item(ByVal key As String) As Object
        ' --- Purpose: Give access to item values by key
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Get
            If IsWhitespaceString(key) Then
                Throw New ArgumentNullException(NameOf(key), "JSON Error: Key cannot be null/empty/whitespace")
            End If
            If Not _data.ContainsKey(key) Then
                Throw New SystemException($"JSON Error: Key not found: {key}")
            End If
            Return _data(key)
        End Get
        Set(value As Object)
            If IsWhitespaceString(key) Then
                Throw New ArgumentNullException(NameOf(key), "JSON Error: Key cannot be null/empty/whitespace")
            End If
            If Not _data.ContainsKey(key) Then
                Throw New SystemException($"JSON Error: Key not found: {key}")
            End If
            If value IsNot Nothing Then
                If value.GetType Is String.Empty.GetType Then
                    value = FromJsonString(CStr(value))
                End If
            End If
            _data(key) = value
        End Set
    End Property

    Public Function ItemOrNull(ByVal key As String) As Object
        ' --- Purpose: Return item value by key, or return Nothing if missing
        ' --- Author : Scott Bakker
        ' --- Created: 09/20/2019
        If IsWhitespaceString(key) Then
            Throw New ArgumentNullException(NameOf(key), "JSON Error: Key cannot be null/empty/whitespace")
        End If
        If Not _data.ContainsKey(key) Then
            Return Nothing
        End If
        Return _data.Item(key)
    End Function

    Public Sub Merge(ByVal jo As JObject)
        ' --- Purpose: Merge a new JObject onto the current one
        ' --- Author : Scott Bakker
        ' --- Created: 09/17/2019
        ' --- Notes  : If any keys are duplicated, the new value overwrites the current value
        If jo Is Nothing OrElse jo.Count = 0 Then
            Exit Sub
        End If
        For Each key As String In jo
            If IsWhitespaceString(key) Then
                Throw New SystemException("JSON Error: Key cannot be null/empty/whitespace")
            End If
            If _data.ContainsKey(key) Then
                ' --- Overwrite current value with new one
                Me(key) = jo(key)
            Else
                Me.Add(key, jo(key))
            End If
        Next
    End Sub

    Public Sub Merge(ByVal dict As Dictionary(Of String, Object))
        ' --- Purpose: Merge a dictionary into the current JObject
        ' --- Author : Scott Bakker
        ' --- Created: 02/11/2020
        ' --- Notes  : If any keys are duplicated, the new value overwrites the current value
        ' ---        : This is processed one key/value at a time to trap errors.
        If dict Is Nothing OrElse dict.Count = 0 Then
            Exit Sub
        End If
        For Each kv As KeyValuePair(Of String, Object) In dict
            If IsWhitespaceString(kv.Key) Then
                Throw New SystemException("JSON Error: Key cannot be null/empty/whitespace")
            End If
            If _data.ContainsKey(kv.Key) Then
                ' --- Overwrite current value with new one
                Me(kv.Key) = kv.Value
            Else
                Me.Add(kv.Key, kv.Value)
            End If
        Next
    End Sub

    Public Sub Remove(ByVal key As String)
        ' --- Purpose: Remove an item from a JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If IsWhitespaceString(key) Then
            Throw New ArgumentNullException(NameOf(key), "JSON Error: Key cannot be null/empty/whitespace")
        End If
        If _data.ContainsKey(key) Then
            _data.Remove(key)
        End If
    End Sub

#Region "ToString"

    Public Overrides Function ToString() As String
        ' --- Purpose: Convert a JObject into a string
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Dim result As New StringBuilder
        result.Append("{")
        Dim addComma As Boolean = False
        For Each kv As KeyValuePair(Of String, Object) In _data
            If addComma Then
                result.Append(",")
            Else
                addComma = True
            End If
            result.Append(ValueToString(kv.Key))
            result.Append(":")
            result.Append(ValueToString(kv.Value))
        Next
        result.Append("}")
        Return result.ToString()
    End Function

    Public Function ToStringSorted() As String
        ' --- Purpose: Sort the keys before returning as a string
        ' --- Author : Scott Bakker
        ' --- Created: 10/17/2019
        Dim result As New StringBuilder
        result.Append("{")
        Dim addComma As Boolean = False
        Dim sorted As New SortedList(_data)
        For i As Integer = 0 To sorted.Count - 1
            If addComma Then
                result.Append(",")
            Else
                addComma = True
            End If
            result.Append(ValueToString(sorted.GetKey(i)))
            result.Append(":")
            result.Append(ValueToString(sorted.GetByIndex(i)))
        Next
        result.Append("}")
        Return result.ToString()
    End Function

    Public Function ToStringFormatted() As String
        ' --- Purpose: Convert this JObject into a string with formatting
        ' --- Author : Scott Bakker
        ' --- Created: 10/17/2019
        Dim indentlevel As Integer = 0
        Return ToStringFormatted(indentlevel)
    End Function

    Friend Function ToStringFormatted(ByRef indentLevel As Integer) As String
        ' --- Purpose: Convert this JObject into a string with formatting
        ' --- Author : Scott Bakker
        ' --- Created: 10/17/2019
        If _data.Count = 0 Then
            Return "{}" ' avoid indent errors
        End If
        Dim result As New StringBuilder
        result.Append("{")
        If indentLevel >= 0 Then
            result.AppendLine()
            indentLevel += 1
        End If
        Dim addComma As Boolean = False
        For Each kv As KeyValuePair(Of String, Object) In _data
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
            result.Append(ValueToString(kv.Key))
            result.Append(":")
            If indentLevel >= 0 Then
                result.Append(" ")
            End If
            result.Append(ValueToString(kv.Value, indentLevel))
        Next
        If indentLevel >= 0 Then
            result.AppendLine()
            If indentLevel > 0 Then
                indentLevel -= 1
            End If
            result.Append(IndentSpace(indentLevel))
        End If
        result.Append("}")
        Return result.ToString()
    End Function

#End Region

#Region "Parse"

    Public Shared Function Parse(ByVal value As String) As JObject
        ' --- Purpose: Convert a string into a JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Dim pos As Integer = 0
        Return Parse(value, pos)
    End Function

    Friend Shared Function Parse(ByVal value As String, ByRef pos As Integer) As JObject
        ' --- Purpose: Convert a partial string into a JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If value Is Nothing OrElse value.Length = 0 Then
            Return Nothing
        End If
        Dim result As New JObject
        Dim tempKey As String
        Dim tempValue As String
        SkipWhitespace(value, pos)
        If value(pos) <> "{" Then
            Throw New SystemException($"JSON Error: Unexpected token to start JObject: {value(pos)}")
        End If
        pos += 1
        Do
            SkipWhitespace(value, pos)
            ' --- check for symbols
            If value(pos) = "}" Then
                pos += 1
                Exit Do ' --- done building JObject
            End If
            If value(pos) = "," Then
                ' --- this logic ignores extra commas, but is ok
                pos += 1
                Continue Do ' --- Next key/value
            End If
            ' --- Get key string
            tempKey = GetToken(value, pos)
            If IsWhitespaceString(tempKey) Then
                Throw New SystemException("JSON Error: Key cannot be null/empty/whitespace")
            End If
            If tempKey.Length <= 2 OrElse
               Not tempKey.StartsWith("""", StringComparison.Ordinal) OrElse
               Not tempKey.EndsWith("""", StringComparison.Ordinal) Then
                Throw New SystemException($"JSON Error: Invalid key format: {tempKey}")
            End If
            ' --- Convert to usable key
            tempKey = JsonValueToObject(tempKey).ToString
            If IsWhitespaceString(tempKey) Then
                Throw New SystemException("JSON Error: Key cannot be null/empty/whitespace")
            End If
            ' --- Check for ":" between key and value
            SkipWhitespace(value, pos)
            If GetToken(value, pos) <> ":" Then
                Throw New SystemException($"JSON Error: Missing colon: {tempKey}")
            End If
            ' --- Check for JObject, JArray
            SkipWhitespace(value, pos)
            If value(pos) = "{" Then
                Dim jo As JObject = Parse(value, pos)
                result.Add(tempKey, jo)
                Continue Do
            End If
            If value(pos) = "[" Then
                Dim ja As JArray = JArray.Parse(value, pos)
                result.Add(tempKey, ja)
                Continue Do
            End If
            tempValue = GetToken(value, pos)
            result.Add(tempKey, JsonValueToObject(tempValue))
        Loop
        Return result
    End Function

#End Region

#Region "Clone"

    Public Shared Function Clone(ByVal jo As JObject) As JObject
        ' --- Purpose: Clones a JObject
        ' --- Author : Scott Bakker
        ' --- Created: 09/20/2019
        Dim result As New JObject
        If jo IsNot Nothing AndAlso jo._data IsNot Nothing Then
            result._data = New Dictionary(Of String, Object)(jo._data)
        End If
        Return result
    End Function

#End Region

End Class
