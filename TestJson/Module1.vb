Imports DA_JsonLibrary

Module Module1

    Sub Main()
        Test1()
        Console.WriteLine()
        Test2()
        Console.WriteLine()
        Console.Write("Press enter to continue...")
        Console.ReadLine()
    End Sub

    Private Sub Test1()
        Try
            Dim jo As New JObject
            jo.Add("ABC", 123)
            jo.Add("abc", 456.789)
            jo.Add("zzzz", 000.000)
            Console.WriteLine(jo.ToString)
            ' ---
            Dim ja As New JArray
            Console.WriteLine(ja.ToString)
            ja.Add(jo)
            ja.Add(Now)
            ja.Add(DateTimeOffset.Now)
            Console.WriteLine(ja.ToString)
            ' ---
            Dim ja3 As New JArray
            ja3.Add(New DateTimeOffset(2020, 2, 1, 8, 15, 23, New TimeSpan(-5, 0, 0)))
            Console.WriteLine(ja3.ToString)
            ' ---
            Dim ja2 As New JArray
            Dim b As Boolean?
            ja2.Add(b)
            b = True
            ja2.Add(b)
            b = False
            ja2.Add(b)
            Console.WriteLine(ja2.ToString)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub Test2()
        Try
            Dim value As String = "{""zzzz"":1,""ABC"":2,""abc"":3}"
            Dim jo As JObject = JObject.Parse(value)
            Console.WriteLine(jo.ToString)
            Console.WriteLine(jo.ToStringSorted)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

End Module
