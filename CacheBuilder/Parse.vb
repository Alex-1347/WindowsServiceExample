Imports System.Text.RegularExpressions

Public Module Proxy

    Public Enum LinkType
        Href = 1
        Src = 2
    End Enum

    Sub ProcessingHTML(ByRef HTML As String)
        Dim HrefRegex As Regex = New Regex("<a\s.*?href=(?:'|"")([^'"">]+)(?:'|"")", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        ProcessingLinks(HTML, HrefRegex, LinkType.Href)
        HrefRegex = Nothing
        Dim LocationRegex As Regex = New Regex("location.href=(?:'|"")([^'"">]+)(?:'|"")", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        ProcessingLinks(HTML, LocationRegex, LinkType.Href)
        LocationRegex = Nothing
        Dim SrcRegex As Regex = New Regex("<img\s.*?src=(?:'|"")([^'"">]+)(?:'|"")", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        ProcessingLinks(HTML, SrcRegex, LinkType.Src)
        SrcRegex = Nothing
        Dim LinkRegex As Regex = New Regex("<link\s.*?href=(?:'|"")([^'"">]+)(?:'|"")", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        ProcessingLinks(HTML, LinkRegex, LinkType.Href)
        LinkRegex = Nothing
        Dim ScriptRegex As Regex = New Regex("<script\s.*?src=(?:'|"")([^'"">]+)(?:'|"")", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        ProcessingLinks(HTML, ScriptRegex, LinkType.Src)
        ScriptRegex = Nothing
    End Sub

    Sub ProcessingLinks(ByRef HTML As String, Regex As Regex, Type As LinkType)
        Dim Links As MatchCollection = Regex.Matches(HTML)
        Dim I As Integer = 0
        While Links.Count > 0
            If Not Links(I).Value.ToLower.Contains("//") Then
                ReplaceOneRelativeLink(HTML, Links(I).Index, Links(I).Value, Type)
                Links = Regex.Matches(HTML)
            End If
            If I < Links.Count - 1 Then
                I += 1
            Else
                Exit While
            End If

        End While
        Links = Nothing
    End Sub

    Sub ReplaceOneRelativeLink(ByRef HTML As String, LinkPosition As Integer, LinkText As String, Type As LinkType)
        Dim Str1 As New Text.StringBuilder()
        Str1.Append(Left(HTML, LinkPosition))               'add left HTML part outside of link
        Dim Pos1 As Integer
        Select Case Type
            Case Type.Href
                Pos1 = InStr(LinkText.ToLower, "href=", CompareMethod.Text)
            Case Type.Src
                Pos1 = InStr(LinkText.ToLower, "src=", CompareMethod.Text)
        End Select
        If Pos1 > 0 Then
            Dim Pos2 = InStr(Pos1 + 1, LinkText.ToLower, """", CompareMethod.Text)
            If Pos2 <= 0 Then
                Pos2 = InStr(Pos1 + 1, LinkText.ToLower, "'", CompareMethod.Text)
            End If
            If Pos2 <= 0 Then
                Debug.Print("Link start not found :" & LinkText)
            Else
                Dim Pos3 As Integer = InStr(Pos2 + 1, LinkText.ToLower, """", CompareMethod.Text)
                If Pos3 <= 0 Then
                    Pos3 = InStr(Pos2 + 1, LinkText.ToLower, "'", CompareMethod.Text)
                End If
                If Pos3 <= 0 Then
                    Debug.Print("Link end not found :" & LinkText)
                Else
                    Dim ClearSiteLink As String = Mid(LinkText, Pos2 + 1, Pos3 - Pos2 - 1)
                    Str1.Append(Left(LinkText, Pos2))           'add left part of link
                    If ClearSiteLink.StartsWith("/") Then
                        Str1.Append(PromocodeCacheCreater.TargetServerRoot)
                        Str1.Append(ClearSiteLink)
                    Else ' link starts with other chars - #, Index.htm, ../
                        Str1.Append(PromocodeCacheCreater.TargetServerPath)
                        Str1.Append("/")
                        Str1.Append(ClearSiteLink)
                    End If
                    If PromocodeCacheCreater.IsUrlCollected Then
                        PromocodeCacheCreater.UrlList.Add(ClearSiteLink)
                    End If
                    Str1.Append(Mid(LinkText, Pos3 + 1))         'add right part of link
                End If
                End If
        Else
            Debug.Print("Link not found : " & LinkText)
        End If
        Str1.Append(Mid(HTML, LinkPosition + Len(LinkText)))    'add right HTML part outside of link
        HTML = Str1.ToString
        Str1 = Nothing
    End Sub

End Module
