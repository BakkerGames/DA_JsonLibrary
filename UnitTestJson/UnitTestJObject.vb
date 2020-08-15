Imports DA_JsonLibrary

<TestClass()>
Public Class UnitTestJObject

    <TestMethod()>
    Public Sub TestNullJObjectDefaultWhitespace()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{}"
        Dim jo As New JObject
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectNullKey()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo(Nothing) = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectNullKeyLevel2()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", Nothing) = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectNullKeyLevel3()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", "key2", Nothing) = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectEmptyKey()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("") = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectEmptyKeyLevel2()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", "") = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectEmptyKeyLevel3()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", "key2", "") = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWhitespaceKey()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("   " & vbTab & vbCrLf) = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWhitespaceKeyLevel2()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", "   " & vbTab & vbCrLf) = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWhitespaceKeyLevel3()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", "key2", "   " & vbTab & vbCrLf) = 123
        Catch ex As ArgumentNullException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

    <TestMethod()>
    Public Sub TestJObjectNullValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":null}"
        Dim jo As New JObject
        jo("key") = Nothing
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectFalseValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":false}"
        Dim jo As New JObject
        jo("key") = False
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectTrueValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":true}"
        Dim jo As New JObject
        jo("key") = True
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectStringValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""abc""}"
        Dim jo As New JObject
        jo("key") = "abc"
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectStringValueCtrlChars()
        ' --- Note: Be careful converting literal strings with "\" between VB and C#
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""\\r\\n\\t\\b\\f\\u00a3\\""}"
        Dim jo As New JObject
        jo("key") = "\r\n\t\b\f\u00a3\"
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectStringActualCtrlChars()
        ' --- Note: Be careful converting literal strings with "\" between VB and C#
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""\r\n\t\b\f\""\u00a3\""\u0000\\""}"
        Dim jo As New JObject
        Dim value As String = $"{vbCr}{vbLf}{vbTab}{vbBack}{vbFormFeed}""£""{Chr(0)}\"
        jo("key") = value
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectStringExpandedChars()
        ' --- Note: In C#, the strings "a\\b" and "a\u005Cb" both collapse to "a\b" and are equal.
        ' ---       In VB.NET, the strings don't collapse and are not equal. This unit test looks
        ' ---       for "AreNotEqual" because it is written in VB.NET. In C#, use "AreEqual".
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""a\\b""}"
        Dim jo As New JObject
        Dim value As String = "a\u005Cb"
        jo("key") = value
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreNotEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIntValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":123}"
        Dim jo As New JObject
        jo("key") = 123
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIntMaxValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":2147483647}"
        Dim jo As New JObject
        jo("key") = Integer.MaxValue
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIntMinValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":-2147483648}"
        Dim jo As New JObject
        jo("key") = Integer.MinValue
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIntMaxPlusOneValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":2147483648}"
        Dim jo As New JObject
        jo("key") = CLng(Integer.MaxValue) + 1
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIntMinMinusOneValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":-2147483649}"
        Dim jo As New JObject
        jo("key") = CLng(Integer.MinValue) - 1
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectLongMaxValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":9223372036854775807}"
        Dim jo As New JObject
        jo("key") = Long.MaxValue
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectLongMinValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":-9223372036854775808}"
        Dim jo As New JObject
        jo("key") = Long.MinValue
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectLongMaxPlusOneValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":9223372036854775808}"
        Dim jo As New JObject
        jo("key") = CDec(Long.MaxValue) + 1
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectLongMinMinusOneValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":-9223372036854775809}"
        Dim jo As New JObject
        jo("key") = CDec(Long.MinValue) - 1
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectPiValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":3.14159265358979}"
        Dim jo As New JObject
        jo("key") = Math.PI
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDoubleValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":123.45}"
        Dim jo As New JObject
        jo("key") = 123.45
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDoubleExponentValue()
        ' arrange
        Dim actualValue As Double
        Dim expectedValue As Double = 1.2345E+50
        Dim jo As JObject = JObject.Parse("{""key"":1.2345e50}")
        ' act
        actualValue = CDbl(jo.Item("key"))
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectToStringSorted()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""abc"":1,""ABC"":-1,""def"":2,""xyz"":99}"
        Dim jo As New JObject
        jo("abc") = 1
        jo("ABC") = -1
        jo("xyz") = 99
        jo("def") = 2
        ' act
        actualValue = jo.ToStringSorted()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDoubleExponentNegativeValue()
        ' arrange
        Dim actualValue As Double
        Dim expectedValue As Double = 1.2345E-50
        Dim jo As JObject = JObject.Parse("{""key"":1.2345e-50}")
        ' act
        actualValue = CDbl(jo.Item("key"))
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDoubleExponentUpperValue()
        ' arrange
        Dim actualValue As Double
        Dim expectedValue As Double = 1.2345E+50
        Dim jo As JObject = JObject.Parse("{""key"":1.2345E50}")
        ' act
        actualValue = CDbl(jo.Item("key"))
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDoubleExponentUpperNegativeValue()
        ' arrange
        Dim actualValue As Double
        Dim expectedValue As Double = 1.2345E-50
        Dim jo As JObject = JObject.Parse("{""key"":1.2345E-50}")
        ' act
        actualValue = CDbl(jo.Item("key"))
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDateValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02""}"
        Dim jo As New JObject
        jo("key") = Date.Parse("01/02/2017")
        ' act
        actualValue = jo.ToString()
        Console.WriteLine(expectedValue)
        Console.WriteLine(actualValue)
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02 16:42:25""}"
        Dim jo As New JObject
        jo("key") = Date.Parse("01/02/2017 16:42:25")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeMillisecondsValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02 16:42:25.123""}"
        Dim jo As New JObject
        jo("key") = Date.Parse("01/02/2017 16:42:25.123")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeOffsetValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02T16:42:25-05:00""}"
        Dim jo As New JObject
        jo("key") = DateTimeOffset.Parse("2017-01-02T16:42:25-05:00")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeOffsetMilliValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02T16:42:25.1234567-05:00""}"
        Dim jo As New JObject
        jo("key") = DateTimeOffset.Parse("2017-01-02T16:42:25.1234567-05:00")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeOffsetValueZ()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02T16:42:25Z""}"
        Dim jo As New JObject
        jo("key") = DateTimeOffset.Parse("2017-01-02T16:42:25Z")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeOffsetMilliValueZ()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02T16:42:25.1234567Z""}"
        Dim jo As New JObject
        jo("key") = DateTimeOffset.Parse("2017-01-02T16:42:25.1234567Z")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDateTimeMultiFormat()
        Dim jo As New JObject
        jo("date") = "2020-02-15"
        jo("time") = "23:59:59"
        jo("timemilli") = "23:59:59.123"
        jo("datetime") = "2020-02-15 23:59:59"
        jo("datetimemilli") = "2020-02-15 23:59:59.123"
        jo("stringdate") = "02/01/2020"
        Dim jo1 As JObject = JObject.Parse(jo.ToString)
        If Not jo1("date").GetType() = GetType(Date) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("time").GetType() = GetType(TimeSpan) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("timemilli").GetType() = GetType(TimeSpan) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("datetime").GetType() = GetType(Date) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("datetimemilli").GetType() = GetType(Date) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("stringdate").GetType() = GetType(Date) Then
            Dim s As String = jo1("stringdate").GetType().ToString
            Assert.IsTrue(False)
            Exit Sub
        End If
        Assert.IsTrue(True)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDateTimeOffsetMultiFormat()
        Dim jo As New JObject
        jo("datetimeoffset") = "2020-02-01T12:23:34-05:00"
        jo("datetimeoffsetmilli") = "2020-02-01T12:23:34.1234567-05:00"
        jo("stringdate") = "02/01/2020"
        Dim jo1 As JObject = JObject.Parse(jo.ToString)
        If Not jo1("datetimeoffset").GetType() = GetType(DateTimeOffset) OrElse
           Not jo1("datetimeoffsetmilli").GetType() = GetType(DateTimeOffset) OrElse
           Not jo1("stringdate").GetType() = GetType(Date) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        Assert.IsTrue(True)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDateFromString()
        Dim s As String = "{""key"":""2020-02-01 12:23:34""}"
        Dim jo As JObject = JObject.Parse(s)
        Dim actualValue As Object
        actualValue = jo("key")
        If actualValue.GetType() = GetType(Date) Then
            Assert.IsTrue(True)
        Else
            Assert.IsTrue(False)
        End If
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDateTimeOffsetFromString()
        Dim s As String = "{""key"":""2020-02-01T12:23:34-05:00""}"
        Dim jo As JObject = JObject.Parse(s)
        Dim actualValue As Object
        actualValue = jo("key")
        If actualValue.GetType() = GetType(DateTimeOffset) Then
            Assert.IsTrue(True)
        Else
            Assert.IsTrue(False)
        End If
    End Sub

    <TestMethod()>
    Public Sub TestJObjectJObjectValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":{""newkey"":456}}"
        Dim jo As New JObject
        Dim jo2 As New JObject()
        jo2("newkey") = 456
        jo("key") = jo2
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectJObjectValueFormatted()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{" + vbCrLf +
                                      "  ""key"": {" + vbCrLf +
                                      "    ""newkey"": 456" + vbCrLf +
                                      "  }" + vbCrLf +
                                      "}"
        Dim jo As New JObject
        Dim jo2 As New JObject()
        jo2("newkey") = 456
        jo("key") = jo2
        ' act
        actualValue = jo.ToStringFormatted()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectJArrayValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":[""newkey"",456]}"
        Dim jo As New JObject
        Dim ja As JArray = New JArray()
        ja.Add("newkey")
        ja.Add(456)
        jo("key") = ja
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectJArrayValueFormatted()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{" + vbCrLf +
                                      "  ""key"": [" + vbCrLf +
                                      "    ""newkey""," + vbCrLf +
                                      "    456" + vbCrLf +
                                      "  ]" + vbCrLf +
                                      "}"
        Dim jo As New JObject
        Dim ja As JArray = New JArray()
        ja.Add("newkey")
        ja.Add(456)
        jo("key") = ja
        ' act
        actualValue = jo.ToStringFormatted()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectMultiValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":123,""otherkey"":789.12}"
        Dim jo As New JObject
        jo("key") = 123
        jo("otherkey") = 789.12
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectNewEmpty()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{}"
        Dim jo As JObject = JObject.Parse(expectedValue)
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectItemArray()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""array"":[1,2,3,4]}"
        Dim ia As Integer() = {1, 2, 3, 4}
        Dim jo As New JObject
        jo("array") = ia
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectItemList()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""list"":[1,2,3,4]}"
        Dim ia As List(Of Integer) = {1, 2, 3, 4}.ToList
        Dim jo As New JObject
        jo("list") = ia
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectClone()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""firstitem"":1,""seconditem"":2,""thirditem"":3,""fourthitem"":4}"
        Dim jo1 As New JObject
        jo1("firstitem") = 1
        jo1("seconditem") = 2
        Dim jo2 As New JObject(jo1)
        jo2("thirditem") = 3
        jo2("fourthitem") = 4
        ' act
        actualValue = jo2.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectMerge()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""firstitem"":1,""seconditem"":2,""thirditem"":3,""fourthitem"":4}"
        Dim jo1 As New JObject
        jo1("firstitem") = 1
        jo1("seconditem") = 2
        Dim jo2 As New JObject
        jo2("thirditem") = 3
        jo2("fourthitem") = 4
        jo1.Merge(jo2)
        ' act
        actualValue = jo1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectMergeWithUpdate()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""firstitem"":3,""seconditem"":2,""fourthitem"":4}"
        Dim jo1 As New JObject()
        jo1("firstitem") = 1
        jo1("seconditem") = 2
        Dim jo2 As New JObject()
        jo2("firstitem") = 3
        jo2("fourthitem") = 4
        jo1.Merge(jo2)
        ' act
        actualValue = jo1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWithComments()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""value""}"
        Dim jo As JObject = JObject.Parse("//comment aaa" & vbCrLf &
                                          "{//comment bbb" & vbCrLf &
                                          """key""/*comment ccc*/:/*comment ddd*/""value""//comment eee" & vbCrLf &
                                          "}//comment xyz")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWithCommentsInValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""value/*value*/abc//def""}"
        Dim jo As JObject = JObject.Parse("//comment aaa" & vbCrLf &
                                          "{//comment bbb" & vbCrLf &
                                          """key""/*comment ccc*/:/*comment ddd*/""value/*value*/abc//def""//comment eee" & vbCrLf &
                                          "}//comment xyz")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWithCommentsNumeric()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":123}"
        Dim jo As JObject = JObject.Parse("//comment aaa" & vbCrLf &
                                          "{//comment bbb" & vbCrLf &
                                          """key""/*comment ccc*/:/*comment ddd*/123//comment eee" & vbCrLf &
                                          "}//comment xyz")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWithCommentsInvalid()
        Try
            JObject.Parse("{""abc"":/123}")
#Disable Warning CA1031 ' Do not catch general exception types
        Catch ex As Exception
            Assert.IsTrue(True, "Expected exception: " & ex.Message)
            Exit Sub
#Enable Warning CA1031 ' Do not catch general exception types
        End Try
        Assert.Fail("Exception not thrown when expected")
    End Sub

    <TestMethod()>
    Public Sub TestJObjectWhitespace()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key1"":""value"",""key2"":1234,""key3"":false}"
        Dim jo As JObject = JObject.Parse(" { ""key1"" : ""value"" , ""key2"" : 1234 , ""key3"" : false } ")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIEnumerable()
        ' arrange
        Dim actualValue As String = ""
        Dim expectedValue As String = "firstitem,seconditem,thirditem,fourthitem,"
        Dim jo1 As New JObject()
        jo1("firstitem") = 1
        jo1("seconditem") = 2
        jo1("thirditem") = 3
        jo1("fourthitem") = 4
        ' act
        For Each s As String In jo1
            actualValue += s + ","
        Next
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectPathSetValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key1"":{""key2"":{""key3"":123}}}"
        Dim jo As New JObject
        ' act
        jo("key1", "key2", "key3") = 123
        actualValue = jo.ToString
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectPathGetValue()
        ' arrange
        Dim actualValue As Integer
        Dim expectedValue As Integer = 123
        Dim jo As New JObject
        ' act
        jo("key1", "key2", "key3") = 123
        actualValue = jo("key1", "key2", "key3")
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectPathGetValueNothing()
        ' arrange
        Dim actualValue As Object
        Dim expectedValue As Object = Nothing
        ' act
        Dim jo As New JObject
        actualValue = jo("key1", "key2", "key3")
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectPathGetValueNothingLevel2()
        ' arrange
        Dim actualValue As Object
        Dim expectedValue As Object = Nothing
        ' act
        Dim jo As New JObject
        jo("key1") = New JObject
        actualValue = jo("key1", "key2", "key3")
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectInvalidTypeException()
        ' arrange
        Dim jo As New JObject
        Try
            ' act
            jo("key1", "key2") = 456
            jo("key1", "key2", "key3") = 123
        Catch ex As ArgumentException
            ' succeeded
            Assert.IsTrue(True)
            Exit Sub
        Catch ex As Exception
            ' failed
            Assert.Fail()
            Exit Sub
        End Try
        Assert.Fail()
    End Sub

End Class
