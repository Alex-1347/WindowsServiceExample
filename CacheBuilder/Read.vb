Imports System.Net
Module Read
    Async Function ReadHtml(URL As String) As Task(Of String)
        Dim Client As New WebClient
        Dim Resp As Byte() = Await Client.DownloadDataTaskAsync(New Uri(URL))
        If Resp.Length = 0 Then
            Client.Dispose()
            Return ""
        End If
        Dim ContentHeaders() As String = Client.ResponseHeaders.GetValues("Content-Type")
        If ContentHeaders(0).StartsWith("text/html") Then
            Dim Ret1 As String = Text.UTF8Encoding.UTF8.GetString(Resp)
            If String.IsNullOrWhiteSpace(Ret1) Then
                Client.Dispose()
                Return ""
            Else
                Client.Dispose()
                Return Ret1
            End If
        Else
            Client.Dispose()
            Return ""
        End If
    End Function
End Module
