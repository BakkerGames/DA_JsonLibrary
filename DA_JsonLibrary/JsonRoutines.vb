﻿' --- Purpose: Provide a set of routines to support JSON Object and JSON Array classes
' --- Author : Scott Bakker
' --- Created: 09/13/2019
' --- LastMod: 08/11/2020

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

#Region "Constants"

    Private Const _dateFormat As String = "yyyy-MM-dd"

    ' --- These are unspecified time zones used for local times only
    Private Const _timeFormat As String = "HH:mm:ss"
    Private Const _timeMilliFormat As String = "HH:mm:ss.fff"
    Private Const _dateTimeFormat As String = "yyyy-MM-dd HH:mm:ss"
    Private Const _dateTimeMilliFormat As String = "yyyy-MM-dd HH:mm:ss.fff"

    ' --- These are precise ISO 8601 format, with the "T" in the 11th position
    Private Const _dateTimeOffsetFormat As String = "yyyy-MM-ddTHH:mm:sszzz"
    Private Const _dateTimeOffsetMilliFormat As String = "yyyy-MM-ddTHH:mm:ss.fffffffzzz"

    ' --- TimeSpan formats are very different
    Private Const _timeSpanFormat As String = "c" ' [-][d'.']hh':'mm':'ss['.'fffffff]

    ' --- ToStringFormatted
    Private Const _indentSpaceSize As Integer = 2

#End Region

#Region "Public Routines"

    Public Shared Function IndentSpace(ByVal indentLevel As Integer) As String
        ' --- Purpose: Return a string with the proper number of spaces or tabs
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        If indentLevel <= 0 Then
            Return ""
        End If
        Return Space(indentLevel * _indentSpaceSize)
    End Function

    Public Shared Function ValueToString(ByVal value As Object) As String
        ' --- Purpose: Return a value in proper JSON string format without indenting
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        Dim indentLevel As Integer = -1 ' --- don't indent
        Return ValueToString(value, indentLevel)
    End Function

#End Region

#Region "Internal Friend routines"

    Protected Friend Shared Function ValueToString(ByVal value As Object, ByRef indentLevel As Integer) As String
        ' --- Purpose: Return a value in proper JSON string format
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        ' --- LastMod: 08/11/2020

        If value Is Nothing Then
            Return "null"
        End If

        ' --- Get the type for comparison
        Dim t As Type = value.GetType

        ' --- Check for generic list types
        If t.IsGenericType Then
            Dim result As New StringBuilder
            result.Append("["c)
            If indentLevel >= 0 Then
                indentLevel += 1
            End If
            Dim addComma As Boolean = False
            For Each obj As Object In CType(value, IEnumerable)
                If addComma Then
                    result.Append(","c)
                Else
                    addComma = True
                End If
                If indentLevel >= 0 Then
                    result.AppendLine()
                End If
                If indentLevel > 0 Then
                    result.Append(IndentSpace(indentLevel))
                End If
                result.Append(ValueToString(obj))
            Next
            If indentLevel >= 0 Then
                result.AppendLine()
                If indentLevel > 0 Then
                    indentLevel -= 1
                End If
                result.Append(IndentSpace(indentLevel))
            End If
            result.Append("]"c)
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
            result.Append("["c)
            If indentLevel >= 0 Then
                indentLevel += 1
            End If
            Dim addComma As Boolean = False
            For i As Integer = 0 To CType(value, Array).Length - 1
                If addComma Then
                    result.Append(","c)
                Else
                    addComma = True
                End If
                If indentLevel >= 0 Then
                    result.AppendLine()
                    result.Append(IndentSpace(indentLevel))
                End If
                Dim obj As Object = CType(value, Array).GetValue(i)
                result.Append(ValueToString(obj, indentLevel))
            Next
            If indentLevel >= 0 Then
                result.AppendLine()
                If indentLevel > 0 Then
                    indentLevel -= 1
                End If
                result.Append(IndentSpace(indentLevel))
            End If
            result.Append("]"c)
            Return result.ToString
        End If

        ' --- Check for individual types

        Select Case t

            Case GetType(String)
                Dim result As New StringBuilder
                result.Append(""""c)
                For Each c As Char In value.ToString
                    result.Append(ToJsonChar(c))
                Next
                result.Append(""""c)
                Return result.ToString

            Case GetType(Char)
                Dim result As New StringBuilder
                result.Append(""""c)
                result.Append(ToJsonChar(CChar(value)))
                result.Append(""""c)
                Return result.ToString

            Case GetType(Guid)
                Return $"""{value}"""

            Case GetType(Boolean)
                If CBool(value) Then
                    Return "true"
                Else
                    Return "false"
                End If

            Case GetType(DateTimeOffset)
                Dim result As String
                If CType(value, DateTimeOffset).Millisecond = 0 Then
                    result = CType(value, DateTimeOffset).ToString(_dateTimeOffsetFormat, Nothing)
                Else
                    result = CType(value, DateTimeOffset).ToString(_dateTimeOffsetMilliFormat, Nothing)
                End If
                If result.EndsWith("+00:00") OrElse result.EndsWith("-00:00") Then
                    result = $"{result.Substring(0, result.Length - 6)}Z"
                End If
                Return $"""{result}"""

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

            Case GetType(TimeSpan)
                Return CType(value, TimeSpan).ToString(_timeSpanFormat)

            Case GetType(JObject)
                Return CType(value, JObject).ToStringFormatted(indentLevel)

            Case GetType(JArray)
                Return CType(value, JArray).ToStringFormatted(indentLevel)

            Case GetType(Single), GetType(Double), GetType(Decimal)
                ' --- Remove trailing decimal zeros. This is not necessary or part
                ' --- of the JSON specification, but it will be impossible to
                ' --- compare two JSON string representations without doing this.
                Return NormalizeDecimal(value.ToString)

            Case GetType(Byte),
                 GetType(SByte),
                 GetType(Short),
                 GetType(Integer),
                 GetType(Long),
                 GetType(UShort),
                 GetType(UInteger),
                 GetType(ULong)
                ' --- Let ToString do all the work
                Return value.ToString

        End Select

        ' --- Unknown ---
        Throw New SystemException($"JSON Error: Unknown object type: {t}")

    End Function

    Private Shared Function NormalizeDecimal(ByVal value As String) As String
        ' --- Purpose: Gets rid of trailing decimal zeros to normalize value
        ' --- Author : Scott Bakker
        ' --- Created: 03/19/2020
        ' --- LastMod: 08/11/2020
        If value.Contains("E") OrElse value.Contains("e") Then
            ' --- Scientific notation, leave alone
            Return value
        End If
        If Not value.Contains(".") Then
            Return value
        End If
        If value.StartsWith(".") Then
            ' --- Cover leading decimal place
            value = "0" + value
        End If
        Return value.TrimEnd("0"c).TrimEnd("."c)
    End Function

    Protected Friend Shared Function FromSerializedString(ByVal value As String) As String
        ' --- Purpose: Convert a string with escaped characters into control codes
        ' --- Author : Scott Bakker
        ' --- Created: 09/17/2019
        ' --- LastMod: 08/11/2020
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
                        result.Append(""""c)
                    Case "\"c
                        result.Append("\"c)
                    Case "/"c
                        result.Append("/"c)
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
        If lastBackslash Then
            Throw New SystemException($"JSON Error: Unexpected trailing backslash")
        End If
        Return result.ToString
    End Function

    Protected Friend Shared Function GetToken(ByVal value As String, ByRef pos As Integer) As String
        ' --- Purpose: Get a single token from string value for parsing
        ' --- Author : Scott Bakker
        ' --- Created: 09/13/2019
        ' --- LastMod: 04/17/2020
        ' Notes  : Does not do escaped character expansion here, just passes exact value.
        '        : Properly handles \" within strings this way, but nothing else.
        If value Is Nothing Then
            Return Nothing
        End If
        Dim c As Char
        ' --- Ignore whitespece before token
        SkipWhitespace(value, pos)
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
                result.Append("\"c)
                result.Append(c)
                lastBackslash = False
            ElseIf inQuote AndAlso c = "\"c Then
                ' --- Remember backslash for next loop, but don't add it to result
                lastBackslash = True
            ElseIf c = """"c Then
                ' --- Check for quotes around a string
                If inQuote Then
                    result.Append(""""c) ' add ending quote
                    Exit Do ' --- Token is done
                End If
                If result.Length > 0 Then
                    ' --- Quote in the middle of a token?
                    Throw New SystemException("JSON Error: Unexpected quote char")
                End If
                result.Append(""""c) ' add beginning quote
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
        ' --- LastMod: 08/11/2020
        If value Is Nothing OrElse value.Length = 0 Then
            Return Nothing
        End If
        Try
            If value.StartsWith("""", StringComparison.Ordinal) AndAlso
               value.EndsWith("""", StringComparison.Ordinal) Then
                value = value.Substring(1, value.Length - 2) ' remove quotes
                If IsTimeSpanValue(value) Then
                    Return TimeSpan.Parse(value)
                End If
                If IsDateTimeOffsetValue(value) Then
                    Return DateTimeOffset.Parse(value)
                End If
                If IsDateTimeValue(value) Then
                    Return Date.Parse(value)
                End If
                ' --- Parse all escaped sequences to chars
                Return FromSerializedString(value)
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
            End If
            If value.Contains(".") Then
                Return CDec(value)
            End If
            If CLng(value) > Integer.MaxValue OrElse CLng(value) < Integer.MinValue Then
                Return CLng(value)
            End If
            Return CInt(value)
        Catch ex As Exception
            Throw New SystemException($"JSON Error: Value not recognized: {value}{vbCrLf}{ex.Message}")
        End Try
    End Function

    Protected Friend Shared Sub SkipBOM(ByVal value As String, ByRef pos As Integer)
        ' --- Purpose: Skip over Byte-Order Mark (BOM) at the beginning of a stream
        ' --- Author : Scott Bakker
        ' --- Created: 05/20/2020
        ' --- LastMod: 08/11/2020
        ' --- UTF-8 BOM = 0xEF,0xBB,0xBF = 239,187,191
        If value Is Nothing Then
            Exit Sub
        End If
        If pos + 3 <= value.Length Then
            If Asc(value(pos)) = 239 AndAlso
               Asc(value(pos + 1)) = 187 AndAlso
               Asc(value(pos + 2)) = 191 Then
                pos += 3 ' --- Move past BOM
                Exit Sub
            End If
        End If
    End Sub

    Protected Friend Shared Sub SkipWhitespace(ByVal value As String, ByRef pos As Integer)
        ' --- Purpose: Skip over any whitespace characters or any recognized comments
        ' --- Author : Scott Bakker
        ' --- Created: 09/23/2019
        ' --- LastMod: 08/11/2020
        ' --- Notes  : Comments consist of "/*...*/" or "//" to eol (aka line comment).
        '            : "//" comments don't need an eol if at the end, but "/*" does need "*/".
        If value Is Nothing Then
            Exit Sub
        End If
        Dim inComment As Boolean = False
        Dim inLineComment As Boolean = False
        While pos < value.Length
            If inComment Then
                If value(pos) = "/"c AndAlso value(pos - 1) = "*"c Then ' --- found ending "*/"
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
                pos += 2 ' skip over "/*"
                Continue While
            End If
            If value(pos) = "/"c AndAlso pos + 1 < value.Length AndAlso value(pos + 1) = "/"c Then
                inLineComment = True
                pos += 2 ' skip over "//"
                Continue While
            End If
            If IsWhitespace(value(pos)) Then
                pos += 1
                Continue While
            End If
            Exit While
        End While
        If inComment Then
            Throw New SystemException($"JSON Error: Comment starting with /* is not terminated by */")
        End If
    End Sub

    Protected Friend Shared Function IsWhitespaceString(ByVal value As String) As Boolean
        ' --- Purpose: Determine if a string contains only whitespace
        ' --- Author : Scott Bakker
        ' --- Created: 02/12/2020
        ' --- LastMod: 04/06/2020
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
        ' --- LastMod: 05/21/2020
        If c = " "c Then Return True
        If c = vbCr Then Return True
        If c = vbLf Then Return True
        If c = vbTab Then Return True
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

    Private Shared Function IsTimeSpanValue(ByVal value As String) As Boolean
        ' --- Purpose: Determine if value converts to a Time without a Date
        ' --- Author : Scott Bakker
        ' --- Created: 04/27/2020
        ' --- LastMod: 08/11/2020
        If String.IsNullOrEmpty(value) Then Return False
        If Not value.Contains(":") Then Return False
        If value.Contains("/") Then Return False
        If value.Substring(1).Contains("-") Then Return False ' allowed as first char
        ' --- Try to convert to a timespan format
        Dim tempTsValue As TimeSpan
        If TimeSpan.TryParse(value, tempTsValue) Then
            Return True
        End If
        Return False
    End Function

    Private Shared Function IsDateTimeValue(ByVal value As String) As Boolean
        ' --- Purpose: Determine if value converts to a DateTime
        ' --- Author : Scott Bakker
        ' --- Created: 02/19/2020
        ' --- LastMod: 08/11/2020
        If String.IsNullOrEmpty(value) Then Return False
        Dim tempValue As Date
        If Date.TryParse(value, tempValue) Then
            Return True
        End If
        Return False
    End Function

    Private Shared Function IsDateTimeOffsetValue(ByVal value As String) As Boolean
        ' --- Purpose: Determine if value converts to a DateTimeOffset
        ' --- Author : Scott Bakker
        ' --- Created: 02/19/2020
        ' --- LastMod: 08/11/2020
        If String.IsNullOrEmpty(value) Then Return False
        ' --- The "T" in the 11th position is used to indicate DateTimeOffset
        If value.Length < 11 OrElse value(10) <> "T"c Then
            Return False
        End If
        Dim tempValue As DateTimeOffset
        If DateTimeOffset.TryParse(value, tempValue) Then
            Return True
        End If
        Return False
    End Function

#End Region

End Class
