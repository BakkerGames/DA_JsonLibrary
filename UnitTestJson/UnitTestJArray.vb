Imports DA_JsonLibrary

<TestClass()>
Public Class UnitTestJArray

    <TestMethod()>
    Public Sub TestNullJArrayNoWhitespace()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[]"
        Dim ja As New JArray
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayNullValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[null]"
        Dim ja As New JArray
        ja.Add(Nothing)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayFalseValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[false]"
        Dim ja As New JArray
        ja.Add(False)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayTrueValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[true]"
        Dim ja As New JArray
        ja.Add(True)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayStringValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[""abc""]"
        Dim ja As New JArray
        ja.Add("abc")
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayIntValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[123]"
        Dim ja As New JArray
        ja.Add(123)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayDoubleValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[123.45]"
        Dim ja As New JArray
        ja.Add(123.45)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayJObjectValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[123.45,{""key"":""value""}]"
        Dim ja As New JArray
        ja.Add(123.45)
        Dim jo As New JObject()
        jo.Add("key", "value")
        ja.Add(jo)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayJObjectValueFormatted()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[" + vbCrLf +
                                      "  123.45," + vbCrLf +
                                      "  {" + vbCrLf +
                                      "    ""key"": ""value""" + vbCrLf +
                                      "  }" + vbCrLf +
                                      "]"
        Dim ja As New JArray
        ja.Add(123.45)
        Dim jo As New JObject()
        jo.Add("key", "value")
        ja.Add(jo)
        ' act
        actualValue = ja.ToStringFormatted()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayJArrayValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[123.45,[""key"",""value""]]"
        Dim ja As New JArray
        ja.Add(123.45)
        Dim ja2 As New JArray
        ja2.Add("key")
        ja2.Add("value")
        ja.Add(ja2)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayJArrayValueFormatted()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[" + vbCrLf +
                                      "  123.45," + vbCrLf +
                                      "  [" + vbCrLf +
                                      "    ""key""," + vbCrLf +
                                      "    ""value""" + vbCrLf +
                                      "  ]" + vbCrLf +
                                      "]"
        Dim ja As New JArray
        ja.Add(123.45)
        Dim ja2 As New JArray
        ja2.Add("key")
        ja2.Add("value")
        ja.Add(ja2)
        ' act
        actualValue = ja.ToStringFormatted()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayMultiValue()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[""abc"",123.45]"
        Dim ja As New JArray
        ja.Add("abc")
        ja.Add(123.45)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayMultiValueNull()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[""abc"",123.45,null]"
        Dim ja As New JArray
        ja.Add("abc")
        ja.Add(123.45)
        ja.Add(Nothing)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayNewEmpty()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[]"
        Dim ja As JArray = JArray.Parse(expectedValue)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayNewValues()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[""abc"",123.45,null]"
        Dim ja As JArray = JArray.Parse(expectedValue)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayItemArray()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[[1,2,3,4]]"
        Dim ia As Integer() = {1, 2, 3, 4}
        Dim ja As New JArray
        ja.Add(ia)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayItemList()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[[1,2,3,4]]"
        Dim ia As Integer() = {1, 2, 3, 4}
        Dim ja As New JArray
        ja.Add(ia)
        ' act
        actualValue = ja.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayAppend()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[1,2,3,4]"
        Dim ja1 As New JArray()
        ja1.Add(1)
        ja1.Add(2)
        Dim ja2 As New JArray()
        ja2.Add(3)
        ja2.Add(4)
        ' act
        ja1.Append(ja2)
        actualValue = ja1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayAppendList()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[1,2,3,4]"
        Dim ja1 As New JArray()
        ja1.Add(1)
        ja1.Add(2)
        Dim ja2 As New List(Of Integer)
        ja2.Add(3)
        ja2.Add(4)
        ' act
        ja1.Append(ja2)
        actualValue = ja1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayAppendArray()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[1,2,3,4]"
        Dim ja1 As New JArray()
        ja1.Add(1)
        ja1.Add(2)
        Dim ja2 As Integer() = {3, 4}
        ' act
        ja1.Append(ja2)
        actualValue = ja1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayWithComments()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[1,2,3,4]"
        Dim ja1 As JArray = JArray.Parse("/*comment*/[/*comment*/1,//comment" & vbCrLf & "2,/*comment*/3,/*comment*/4/*comment*/]//comment")
        ' act
        actualValue = ja1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

    <TestMethod()>
    Public Sub TestJArrayWhitespace()
        ' arrange
        Dim actualValue As String
        Dim expectedValue As String = "[1,2,3,4]"
        Dim ja1 As JArray = JArray.Parse(" [ 1 , 2 , 3 , 4 ] ")
        ' act
        actualValue = ja1.ToString()
        ' assert
        Assert.AreEqual(expectedValue, actualValue)
    End Sub

End Class