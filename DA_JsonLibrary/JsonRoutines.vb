' --- Purpose: Provide a set of routines to support JSON Object and JSON Array classes
' --- Author : Scott Bakker
' --- Created: 09/13/2019

' --- Notes  : DateTime and DateTimeOffset are stored in JObject and JArray properly
'              as objects of those types.
'            : When JObject/JArray are converted to a string, the formats below are
'              used depending on the value type and contents.
'            : However, when converting back from a string, any value which passes
'              the IsDateTimeValue() or IsDateTimeOffsetValue() check will be
'              converted to a DateTime or DateTimeOffset, even if it had only been
'              a string value before. This could have unanticipated consequences.
'              Be careful storing strings which look like dates.

Imports System.Text

Public NotInheritable Class JsonRoutines

    Private Const _dateFormat As String = "yyyy-MM-dd"
    Private Const _timeFormat As String = "HH:mm:ss"
    Private Const _timeMilliFormat As String = "HH:mm:ss.fff"
    Private Const _dateTimeFormat As String = "yyyy-MM-dd HH:mm:ss"
    Private Const _dateTimeMilliFormat As String = "yyyy-MM-dd HH:mm:ss.fff"
    Private Const _dateTimeOffsetFormat As String = "yyyy-MM-dd HH:mm:sszzz"
    Private Const _dateTimeOffsetMilliFormat As String = "yyyy-MM-dd HH:mm:ss.fffffffzzz"

    Friend Const _indentSpaceSize As Integer = 2

    Public Shared Function ValueToString(ByVal value As Object) As String
        ' --- Purpose: Return a value in proper JSON string format without indenting
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Dim indentLevel As Integer = -1 ' --- don't indent
        Return ValueToString(value, indentLevel)
    End Function

#Region "Internal Friend routines"

    Protected Friend Shared Function ValueToString(ByVal value As Object, ByRef indentLevel As Integer) As String
        ' --- Purpose: Return a value in proper JSON string format
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019

        If value Is Nothing Then
            Return "null"
        End If

        ' --- Get the type for comparison
        Dim t As Type = value.GetType

        ' --- Check for generic list types
        If t.IsGenericType Then
            Dim result As New StringBuilder
            result.Append("[")
            If indentLevel >= 0 Then
                indentLevel += 1
            End If
            Dim addComma As String = Nothing
            For Each obj As Object In CType(value, IEnumerable)
                result.Append(addComma)
                If addComma Is Nothing Then addComma = ","
                If indentLevel >= 0 Then
                    result.AppendLine()
                End If
                If indentLevel > 0 Then
                    result.Append(Space(indentLevel * _indentSpaceSize))
                End If
                result.Append(ValueToString(obj))
            Next
            If indentLevel >= 0 Then
                result.AppendLine()
            End If
            If indentLevel > 0 Then
                indentLevel -= 1
            End If
            If indentLevel > 0 Then
                result.Append(Space(indentLevel * _indentSpaceSize))
            End If
            result.Append("]")
            Return result.ToString
        End If

        ' --- Check for byte array, return as hex string "0x00..." with quotes
        If t.IsArray AndAlso t Is GetType(Byte()) Then
            Dim result As New StringBuilder
            result.Append("""0x")
            For Each b As Byte In CType(value, Byte())
                result.Append(b.ToString("x2", Nothing)) ' lowercase hex byte
            Next
            result.Append("""")
            Return result.ToString
        End If

        ' --- Check for array, return in JArray format
        If t.IsArray Then
            Dim result As New StringBuilder
            result.Append("[")
            Dim addComma As String = Nothing
            If indentLevel >= 0 Then
                indentLevel += 1
            End If
            For i As Integer = 0 To CType(value, Array).Length - 1
                Dim obj As Object = CType(value, Array).GetValue(i)
                result.Append(addComma)
                If addComma Is Nothing Then addComma = ","
                If indentLevel >= 0 Then
                    result.AppendLine()
                End If
                If indentLevel > 0 Then
                    result.Append(Space(indentLevel * _indentSpaceSize))
                End If
                result.Append(ValueToString(obj, indentLevel))
            Next
            If indentLevel >= 0 Then
                result.AppendLine()
            End If
            If indentLevel > 0 Then
                indentLevel -= 1
            End If
            If indentLevel > 0 Then
                result.Append(Space(indentLevel * _indentSpaceSize))
            End If
            result.Append("]")
            Return result.ToString
        End If

        ' --- Check for individual types
        Select Case t

            Case GetType(String), GetType(Char)
                Dim result As New StringBuilder
                result.Append("""")
                For Each c As Char In value.ToString
                    result.Append(ToJsonChar(c))
                Next
                result.Append("""")
                Return result.ToString

            Case GetType(Guid)
                Return $"""{value.ToString}"""

            Case GetType(Boolean)
                If CBool(value) Then
                    Return "true"
                Else
                    Return "false"
                End If

            Case GetType(DateTimeOffset)
                If CType(value, DateTimeOffset).Millisecond = 0 Then
                    Return $"""{CType(value, DateTimeOffset).ToString(_dateTimeOffsetFormat, Nothing)}"""
                Else
                    Return $"""{CType(value, DateTimeOffset).ToString(_dateTimeOffsetMilliFormat, Nothing)}"""
                End If

            Case GetType(Date)
                Dim d As Date = CDate(value)
                If d.Hour + d.Minute + d.Second + d.Millisecond = 0 Then
                    Return $"""{d.ToString(_dateFormat, Nothing)}"""
                End If
                If d.Year + d.Month + d.Day = 0 Then
                    If d.Millisecond = 0 Then
                        Return $"""{d.ToString(_timeFormat, Nothing)}"""
                    Else
                        Return $"""{d.ToString(_timeMilliFormat, Nothing)}"""
                    End If
                End If
                If d.Millisecond = 0 Then
                    Return $"""{d.ToString(_dateTimeFormat, Nothing)}"""
                Else
                    Return $"""{d.ToString(_dateTimeMilliFormat, Nothing)}"""
                End If

            Case GetType(JObject)
                Return CType(value, JObject).ToStringFormatted(indentLevel)

            Case GetType(JArray)
                Return CType(value, JArray).ToStringFormatted(indentLevel)

            Case GetType(Byte), GetType(SByte), GetType(Short), GetType(Integer),
                 GetType(Long), GetType(UShort), GetType(UInteger), GetType(ULong),
                 GetType(Single), GetType(Double), GetType(Decimal)
                ' --- Let ToString do all the work
                Return value.ToString

        End Select

        ' --- Unknown ---
        Throw New SystemException($"JSON Error: Unknown object type: {t.ToString}")

    End Function

    Protected Friend Shared Function FromJsonString(ByVal value As String) As String
        ' --- Purpose: Convert a string with escaped characters into control codes
        ' --- Author : Scott Bakker
        ' --- Created: 09/17/2019
        If value Is Nothing Then
            Return Nothing
        End If
        If Not value.Contains("\") Then
            Return value
        End If
        Dim result As New StringBuilder
        Dim lastBackslash As Boolean = False
        Dim unicodeCharCount As Integer = 0
        Dim unicodeValue As String = ""
        For Each c As Char In value
            If unicodeCharCount > 0 Then
                unicodeValue += c
                unicodeCharCount -= 1
                If unicodeCharCount = 0 Then
                    result.Append(Convert.ToChar(Convert.ToUInt16(unicodeValue, 16)))
                    unicodeValue = ""
                End If
            ElseIf lastBackslash Then
                Select Case c
                    Case """"c
                        result.Append("""")
                    Case "\"c
                        result.Append("\")
                    Case "/"c
                        result.Append("/")
                    Case "r"c
                        result.Append(vbCr)
                    Case "n"c
                        result.Append(vbLf)
                    Case "t"c
                        result.Append(vbTab)
                    Case "b"c
                        result.Append(vbBack)
                    Case "f"c
                        result.Append(vbFormFeed)
                    Case "u"c
                        unicodeCharCount = 4
                        unicodeValue = ""
                    Case Else
                        Throw New SystemException($"JSON Error: Unexpected escaped char: {c}")
                End Select
                lastBackslash = False
            ElseIf c = "\"c Then
                lastBackslash = True
            Else
                result.Append(c)
            End If
        Next
        Return result.ToString
    End Function

    Protected Friend Shared Function GetToken(ByRef pos As Integer, ByVal value As String) As String
        ' Purpose: Get a single token from string value for parsing
        ' Author : Scott Bakker
        ' Created: 09/13/2019
        ' Notes  : Does not do escaped character expansion here, just passes exact value.
        '        : Properly handles \" within strings properly this way, but nothing else.
        If value Is Nothing Then
            Return Nothing
        End If
        Dim c As Char
        ' --- Ignore whitespece before token
        SkipWhitespace(pos, value)
        ' --- Get first char, check for special symbols
        c = value(pos)
        pos += 1
        ' --- Stop if one-character JSON symbol found
        If IsJsonSymbol(c) Then
            Return c
        End If
        ' --- Have to build token char by char
        Dim result As New StringBuilder
        Dim inQuote As Boolean = False
        Dim lastBackslash As Boolean = False
        Do
            ' --- Check for whitespace or symbols to end token
            If Not inQuote Then
                If IsWhitespace(c) Then
                    Exit Do
                End If
                If IsJsonSymbol(c) Then
                    pos -= 1 ' move back one char so symbol can be read next time
                    Exit Do
                End If
                ' --- Any comments end the token
                If c = "/"c AndAlso pos < value.Length Then
                    If value(pos) = "*"c OrElse value(pos) = "/"c Then
                        pos -= 1 ' move back one char so comment can be read next time
                        Exit Do
                    End If
                End If
                If c <> """"c AndAlso Not IsJsonValueChar(c) Then
                    Throw New SystemException($"JSON Error: Unexpected character: {c}")
                End If
            End If
            ' --- Check for escaped chars
            If inQuote AndAlso lastBackslash Then
                ' --- Add backslash and character, no expansion here
                result.Append("\")
                result.Append(c)
                lastBackslash = False
            ElseIf inQuote AndAlso c = "\"c Then
                ' --- Remember backslash for next loop, but don't add it to result
                lastBackslash = True
            ElseIf c = """"c Then
                ' --- Check for quotes around a string
                If inQuote Then
                    result.Append("""") ' add ending quote
                    Exit Do ' --- Token is done
                End If
                If result.Length > 0 Then
                    ' --- Quote in the middle of a token?
                    Throw New SystemException("JSON Error: Unexpected quote char")
                End If
                result.Append("""") ' add beginning quote
                inQuote = True
            Else
                ' --- Add this char
                result.Append(c)
            End If
            ' --- Get char for next loop
            c = value(pos)
            pos += 1
        Loop While pos <= value.Length
        Return result.ToString
    End Function

    Protected Friend Shared Function JsonValueToObject(ByVal value As String) As Object
        ' --- Purpose: Convert a string representation of a value to an actual object
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If value Is Nothing OrElse value.Length = 0 Then
            Return Nothing
        End If
        Try
            If value.StartsWith("""", StringComparison.Ordinal) AndAlso value.EndsWith("""", StringComparison.Ordinal) Then
                value = value.Substring(1, value.Length - 2) ' remove quotes
                If IsDateTimeValue(value) Then
                    Return CDate(value)
                ElseIf IsDateTimeOffsetValue(value) Then
                    Return CType(value, DateTimeOffset)
                End If
                ' --- Parse all escaped sequences to chars
                Return FromJsonString(value)
            End If
            If value = "null" Then
                Return Nothing
            End If
            If value = "true" Then
                Return True
            End If
            If value = "false" Then
                Return False
            End If
            ' --- must be numeric
            If value.Contains("e") OrElse value.Contains("E") Then
                Return CDbl(value)
            ElseIf value.Contains(".") Then
                Return CDec(value)
            ElseIf CLng(value) > Integer.MaxValue OrElse CLng(value) < Integer.MinValue Then
                Return CLng(value)
            Else
                Return CInt(value)
            End If
        Catch ex As Exception
            Throw New SystemException($"JSON Error: Value not recognized: {value}{vbCrLf}{ex.Message}")
        End Try
    End Function

    Protected Friend Shared Sub SkipWhitespace(ByRef pos As Integer, ByVal value As String)
        ' Purpose: Skip over any whitespace characters or any recognized comments
        ' Author : Scott Bakker
        ' Created: 09/23/2019
        ' Notes  : Comments consist of /*...*/ or // to eol (aka line comment)
        '        : An unterminated comment is not an error, it is just all skipped
        If value Is Nothing Then
            Exit Sub
        End If
        Dim inComment As Boolean = False
        Dim inLineComment As Boolean = False
        While pos < value.Length
            If inComment Then
                If value(pos) = "/"c AndAlso value(pos - 1) = "*"c Then ' found ending */
                    inComment = False
                End If
                pos += 1
                Continue While
            End If
            If inLineComment Then
                If value(pos) = vbCr OrElse value(pos) = vbLf Then
                    inLineComment = False
                End If
                pos += 1
                Continue While
            End If
            If value(pos) = "/"c AndAlso pos + 1 < value.Length AndAlso value(pos + 1) = "*"c Then
                inComment = True
                pos += 3 ' must be sure to skip enough so /*/ pattern doesn't work but /**/ does
                Continue While
            End If
            If value(pos) = "/"c AndAlso pos + 1 < value.Length AndAlso value(pos + 1) = "/"c Then
                inLineComment = True
                pos += 2 ' skip over //
                Continue While
            End If
            If IsWhitespace(value(pos)) Then
                pos += 1
                Continue While
            End If
            Exit While
        End While
    End Sub

    Protected Friend Shared Function IsWhitespaceString(ByVal value As String) As Boolean
        ' --- Purpose: Determine if a string contains only whitespace
        ' --- Author : Scott Bakker
        ' --- Created: 02/12/2020
        If value Is Nothing Then
            Return True
        End If
        If value.Length = 0 Then
            Return True
        End If
        For Each c As Char In value
            If Not IsWhitespace(c) Then
                Return False
            End If
        Next
        Return True
    End Function

#End Region

#Region "Private routines"

    Private Shared Function ToJsonChar(ByVal c As Char) As String
        ' --- Purpose: Return a character in proper JSON format
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If c = "\"c Then Return "\\"
        If c = """"c Then Return "\"""
        If c = CChar(vbCr) Then Return "\r"
        If c = CChar(vbLf) Then Return "\n"
        If c = CChar(vbTab) Then Return "\t"
        If c = CChar(vbBack) Then Return "\b"
        If c = CChar(vbFormFeed) Then Return "\f"
        Dim cAsc As Integer = AscW(c)
        If cAsc < 32 OrElse cAsc >= 127 Then
            Return "\u" & cAsc.ToString("x4", Nothing) ' always lowercase
        End If
        Return c
    End Function

    Private Shared Function IsWhitespace(ByVal c As Char) As Boolean
        ' --- Purpose: Check for recognized whitespace characters
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If c = " "c Then Return True
        If c = vbCr Then Return True
        If c = vbLf Then Return True
        If c = vbTab Then Return True
        If c = vbBack Then Return True
        If c = vbFormFeed Then Return True
        Return False
    End Function

    Private Shared Function IsJsonSymbol(ByVal c As Char) As Boolean
        ' --- Purpose: Check for recognized JSON symbol chars which are tokens by themselves
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If c = "{"c Then Return True
        If c = "}"c Then Return True
        If c = "["c Then Return True
        If c = "]"c Then Return True
        If c = ":"c Then Return True
        If c = ","c Then Return True
        Return False
    End Function

    Private Shared Function IsJsonValueChar(ByVal c As Char) As Boolean
        ' --- Purpose: Check for any valid characters in a non-string value
        ' --- Author : Scott Bakker
        ' --- Created: 09/23/2019
        Select Case c
            ' --- null
            Case "n"c : Return True
            Case "u"c : Return True
            Case "l"c : Return True
            ' --- true
            Case "t"c : Return True
            Case "r"c : Return True
            Case "e"c : Return True
            ' --- false
            Case "f"c : Return True
            Case "a"c : Return True
            Case "s"c : Return True
            ' --- numeric value
            Case "0"c To "9"c : Return True
            Case "-"c : Return True
            Case "."c : Return True
            Case "E"c : Return True ' also "e", checked for above
        End Select
        Return False
    End Function

    Private Shared Function IsDateTimeValue(ByVal value As String) As Boolean
        ' --- Purpose: Determine if value converts to a DateTimeOffset
        ' --- Author : Scott Bakker
        ' --- Created: 02/19/2020
        If String.IsNullOrEmpty(value) Then Return False
        Dim len As Integer = value.Length
        If len <> 8 AndAlso len <> 10 AndAlso len <> 12 AndAlso len <> 19 AndAlso len <> 23 Then
            Return False
        End If
        Try
            If Not IsDate(value) Then Return False
            If len = 8 Then
                Return (CDate(value).ToString(_timeFormat) = value)
            ElseIf len = 10 Then
                Return (CDate(value).ToString(_dateFormat) = value)
            ElseIf len = 12 Then
                Return (CDate(value).ToString(_timeMilliFormat) = value)
            ElseIf len = 19 Then
                Return (CDate(value).ToString(_dateTimeFormat) = value)
            ElseIf len = 23 Then
                Return (CDate(value).ToString(_dateTimeMilliFormat) = value)
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Shared Function IsDateTimeOffsetValue(ByVal value As String) As Boolean
        ' --- Purpose: Determine if value converts to a DateTimeOffset
        ' --- Author : Scott Bakker
        ' --- Created: 02/19/2020
        If String.IsNullOrEmpty(value) Then Return False
        Dim len As Integer = value.Length
        If len <> 25 AndAlso len <> 33 Then
            Return False
        End If
        Try
            If Not IsDate(value) Then Return False
            If len = 25 Then
                Return (CType(value, DateTimeOffset).ToString(_dateTimeOffsetFormat) = value)
            ElseIf len = 33 Then
                Return (CType(value, DateTimeOffset).ToString(_dateTimeOffsetMilliFormat) = value)
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

End Class
