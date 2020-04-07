﻿Imports DA_JsonLibrary

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
            jo.Add(Nothing, 123)
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
            jo.Add("", 123)
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
            jo.Add("   " & vbTab & vbCrLf, 123)
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
        jo.Add("key", Nothing)
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
        jo.Add("key", False)
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
        jo.Add("key", True)
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
        jo.Add("key", "abc")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectStringValueCtrlChars()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""\\r\\n\\t\\b\\f\\u00a3""}"
        Dim jo As New JObject
        jo.Add("key", "\r\n\t\b\f\u00a3")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectStringActualCtrlChars()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""\r\n\t\b\f\""\u00a3\""\u0000\\""}"
        Dim jo As New JObject
        jo.Add("key", $"{vbCr}{vbLf}{vbTab}{vbBack}{vbFormFeed}""£""{Chr(0)}\")
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectIntValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":123}"
        Dim jo As New JObject
        jo.Add("key", 123)
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
        jo.Add("key", Integer.MaxValue)
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
        jo.Add("key", Integer.MinValue)
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
        jo.Add("key", CLng(Integer.MaxValue) + 1)
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
        jo.Add("key", CLng(Integer.MinValue) - 1)
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
        jo.Add("key", Long.MaxValue)
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
        jo.Add("key", Long.MinValue)
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
        jo.Add("key", CDec(Long.MaxValue) + 1)
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
        jo.Add("key", CDec(Long.MinValue) - 1)
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
        jo.Add("key", Math.PI)
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
        jo.Add("key", 123.45)
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
        jo.Add("abc", 1)
        jo.Add("ABC", -1)
        jo.Add("xyz", 99)
        jo.Add("def", 2)
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
        jo.Add("key", Date.Parse("01/02/2017"))
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
        jo.Add("key", Date.Parse("01/02/2017 16:42:25"))
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
        jo.Add("key", Date.Parse("01/02/2017 16:42:25.123"))
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeOffsetValue()
        ' Note: Works in Eastern US time zone, others will fail
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02T16:42:25-05:00""}"
        Dim jo As New JObject
        jo.Add("key", DateTimeOffset.Parse("2017-01-02T16:42:25-05:00"))
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDatetimeOffsetMilliValue()
        ' Note: Works in Eastern US time zone, others will fail
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "{""key"":""2017-01-02T16:42:25.1234567-05:00""}"
        Dim jo As New JObject
        jo.Add("key", DateTimeOffset.Parse("2017-01-02T16:42:25.1234567-05:00"))
        ' act
        actualValue = jo.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJObjectDateTimeMultiFormat()
        Dim jo As New JObject
        jo.Add("date", "2020-02-15")
        jo.Add("time", "23:59:59")
        jo.Add("timemilli", "23:59:59.123")
        jo.Add("datetime", "2020-02-15 23:59:59")
        jo.Add("datetimemilli", "2020-02-15 23:59:59.123")
        jo.Add("stringdate", "02/01/2020")
        Dim jo1 As JObject = JObject.Parse(jo.ToString)
        If Not jo1("date").GetType() = GetType(Date) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("time").GetType() = GetType(String) Then
            Assert.IsTrue(False)
            Exit Sub
        End If
        If Not jo1("timemilli").GetType() = GetType(String) Then
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
        jo.Add("datetimeoffset", "2020-02-01T12:23:34-05:00")
        jo.Add("datetimeoffsetmilli", "2020-02-01T12:23:34.1234567-05:00")
        jo.Add("stringdate", "02/01/2020")
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
        jo2.Add("newkey", 456)
        jo.Add("key", jo2)
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
        jo2.Add("newkey", 456)
        jo.Add("key", jo2)
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
        jo.Add("key", ja)
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
        jo.Add("key", ja)
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
        jo.Add("key", 123)
        jo.Add("otherkey", 789.12)
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
        jo.Add("array", ia)
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
        jo.Add("list", ia)
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
        jo1.Add("firstitem", 1)
        jo1.Add("seconditem", 2)
        Dim jo2 As JObject = JObject.Clone(jo1)
        jo2.Add("thirditem", 3)
        jo2.Add("fourthitem", 4)
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
        jo1.Add("firstitem", 1)
        jo1.Add("seconditem", 2)
        Dim jo2 As New JObject
        jo2.Add("thirditem", 3)
        jo2.Add("fourthitem", 4)
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
        jo1.Add("firstitem", 1)
        jo1.Add("seconditem", 2)
        Dim jo2 As New JObject()
        jo2.Add("firstitem", 3)
        jo2.Add("fourthitem", 4)
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
        jo1.Add("firstitem", 1)
        jo1.Add("seconditem", 2)
        jo1.Add("thirditem", 3)
        jo1.Add("fourthitem", 4)
        ' act
        For Each s As String In jo1
            actualValue += s + ","
        Next
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

End Class