Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Runtime.ConstrainedExecution
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp

Module ModMain
    Dim _SW As StreamWriter = Nothing
    Dim _baseURL As String = ConfigurationManager.AppSettings("baseURL")
    Sub Main()
        Dim client = New RestClient(_baseURL)
        Dim rawJson As String = String.Empty
        Dim request As RestRequest = Nothing
        Dim response As RestResponse = Nothing
        Dim responseContent As String = String.Empty
        Dim parseJSON As JObject = Nothing
        Dim car As clsCar = Nothing
        Dim CarryOn As Boolean = False
        Dim LogFileName As String = CurDir() + "\RestAPITest.log"
        Dim CarCount As Integer = 0
        Dim NewCarID As Integer = 0

        '
        ' Open a log file
        '
        _SW = New StreamWriter(LogFileName)

        WriteLog("Start of Test")
        WriteLog("Using " + _baseURL)
        WriteLog("")

        '
        ' Get a list of cars
        '
        CarryOn = (GetCars() > 0)

        If CarryOn Then
            '
            ' Now add a new car
            '
            car = New clsCar
            car.build = 2020
            car.manufacturer = "Vauxall"
            car.model = "Astra"

            '
            ' Seralise this and post to the API
            '
            rawJson = JsonConvert.SerializeObject(car, Formatting.Indented)

            WriteLog("Using /cars endpoint to post a new car using the API")
            request = New RestRequest()
            request.Resource = "/cars"
            request.Method = Method.Post
            request.AddJsonBody(rawJson)
            response = client.Execute(request)
            WriteLog("Status Code : " + response.StatusCode.ToString)
            responseContent = response.Content
            WriteLog("Content : " + responseContent)
            WriteLog("")

            CarryOn = (response.StatusCode.ToString = "OK")

        End If

        '
        ' Get the list of cars again - should see the new car added
        '
        If CarryOn Then
            CarryOn = (GetCars() > 0)
        End If

        If CarryOn Then

            '
            ' Amend the Car ID 1
            '
            car = New clsCar
            car.id = 1
            car.build = "2010"
            car.manufacturer = "Vauxall"
            car.model = "Vectra"

            '
            ' Seralise this and post to the API
            '
            rawJson = JsonConvert.SerializeObject(car, Formatting.Indented)
            WriteLog("Using /cars endpoint to put a change to car ID 1 using the API")
            request = New RestRequest()
            request.Resource = "/cars"
            request.Method = Method.Put
            request.AddJsonBody(rawJson)
            response = client.Execute(request)
            WriteLog("Status Code : " + response.StatusCode.ToString)
            responseContent = response.Content
            WriteLog("Content : " + responseContent)
            WriteLog("")

            CarryOn = (response.StatusCode.ToString = "OK")
        End If

        '
        ' Get the list of cars again - should see car 1 amended
        '
        If CarryOn Then
            CarryOn = (GetCars() > 0)
        End If

        If CarryOn Then

            '
            ' Reset the Car ID 1 back to Ford Model T from 1927
            '
            car = New clsCar
            car.id = 1
            car.build = 1927
            car.manufacturer = "Ford"
            car.model = "Model T"

            '
            ' Seralise this and post to the API
            '
            rawJson = JsonConvert.SerializeObject(car, Formatting.Indented)
            WriteLog("Using /cars endpoint to put a change to car ID 1 using the API")
            request = New RestRequest()
            request.Resource = "/cars"
            request.Method = Method.Put
            request.AddJsonBody(rawJson)
            response = client.Execute(request)
            WriteLog("Status Code : " + response.StatusCode.ToString)
            responseContent = response.Content
            WriteLog("Content : " + responseContent)
            WriteLog("")

            CarryOn = (response.StatusCode.ToString = "OK")
        End If

        '
        ' Get the list of cars again - should see car 1 reset
        '
        If CarryOn Then
            NewCarID = GetCars()
            CarryOn = (NewCarID > 3)
        End If

        '
        ' Delete the car that was added
        '
        If CarryOn Then

            WriteLog("Using /cars/" + NewCarID.ToString + " to delete car ID " + NewCarID.ToString + " using the API")
            request = New RestRequest()
            request.Resource = "/cars/" + NewCarID.ToString
            request.Method = Method.Delete
            response = client.Execute(request)
            WriteLog("Status Code : " + response.StatusCode.ToString)
            WriteLog("")
            CarryOn = (response.StatusCode.ToString = "OK")

            '
            ' Get the list of cars again - should see only 3 cars now
            '
            If CarryOn Then
                GetCars()
            End If
        End If


        WriteLog("")
        WriteLog("End of Job")


        _SW.Close()
        _SW.Dispose()
    End Sub
    Private Function GetCars() As Integer
        Dim ReturnValue As Integer = 0
        Dim client = New RestClient(_baseURL)
        Dim rawJson As String = String.Empty
        Dim request As RestRequest = Nothing
        Dim response As RestResponse = Nothing
        Dim responseContent As String = String.Empty
        Dim parseJSON As JObject = Nothing
        Dim car As clsCar = Nothing
        Dim JSONLines() As String

        WriteLog("Using /cars endpoint to get list of cars from API")
        request = New RestRequest()
        request.Resource = "/cars"
        request.Method = Method.Get
        response = client.Execute(request)
        WriteLog("Status Code : " + response.StatusCode.ToString)
        responseContent = response.Content
        WriteLog("Content : " + responseContent)
        WriteLog("")

        If response.StatusCode.ToString = "OK" Then
            ReturnValue = True
        End If

        If ReturnValue Then
            '
            ' Remove the [ and ]
            '
            responseContent = responseContent.Replace("[", "").Replace("]", "")

            '
            ' As this is a list of JSON objects, put these into JSONLines array so each object can be read
            '
            JSONLines = ParseContent(responseContent)

            If JSONLines.Length > 0 Then

                '
                ' Read each line in JSONLines array, and deserialise into the clsCar object
                '
                WriteLog("")
                WriteLog("Getting the list of card")
                WriteLog("")

                For i As Integer = 0 To JSONLines.Length - 1
                    rawJson = JSONLines(i)
                    If rawJson <> String.Empty Then
                        parseJSON = JObject.Parse(rawJson)
                        car = JsonConvert.DeserializeObject(Of clsCar)(rawJson)
                        WriteLog("ID ......... : " + car.id.ToString)
                        WriteLog("Build ...... : " + car.build.ToString)
                        WriteLog("Manufacturer : " + car.manufacturer)
                        WriteLog("Model ...... : " + car.model)
                        WriteLog("")
                        ReturnValue = car.id
                    End If
                Next
            End If
        End If

        responseContent = Nothing
        response = Nothing
        request = Nothing

        Return ReturnValue
    End Function
    Private Function ParseContent(ByVal inContent As String) As String()
        Dim s As String = String.Empty
        Dim i As Integer = 0
        Dim c As String = String.Empty
        Dim ReturnValue(0) As String
        Dim el As Integer = -1
        Dim jsonStart As Boolean = False

        For i = 0 To inContent.Length - 1
            c = inContent.Substring(i, 1)

            If c = "{" Then
                jsonStart = True
            End If

            If c = "}" Then
                s += c
                el += 1
                If el >= 0 Then
                    ReDim Preserve ReturnValue(el)
                End If
                ReturnValue(el) = s
                s = String.Empty
                jsonStart = False
            End If
            If c = "," Then
                If jsonStart Then
                    s += c
                End If
            Else
                If jsonStart Then
                    s += c
                End If
            End If
        Next

        el += 1
        If el >= 0 Then
            ReDim Preserve ReturnValue(el)
        End If
        ReturnValue(el) = s
        Return ReturnValue
    End Function

    Private Sub WriteLog(ByVal inString As String)
        _SW.WriteLine(DateTime.Now.ToString + " " + inString)
        _SW.Flush()
        Console.WriteLine(inString)
    End Sub
End Module
