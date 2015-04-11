Imports System
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Public Class Form1
    Dim enteredUrl As String
    Dim DownloadUrl() As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        enteredUrl = TextBox1.Text
        'Dim sourceString As String = New System.Net.WebClient().DownloadString(enteredUrl)
        'MsgBox(sourceString)
        'WebBrowser1.Navigate(enteredUrl)
        'Dim htmlSourceCode As String = WebBrowser1.DocumentText.ToString()
        'Dim bodyContent As String = WebBrowser1.Document.Body.InnerText
        Dim request As HttpWebRequest = WebRequest.Create(enteredUrl)
        request.UserAgent = ".NET Framework Test Client"
        Dim response As HttpWebResponse = request.GetResponse()
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
        'Dim str As String = reader.ReadLine()
        Dim str As String = reader.ReadToEnd
        'RichTextBox1.Text = str
        'Dim myRegExp As RegExp
        'Dim myMatches As MatchCollection
        'Dim myMatch As Match
        'myRegExp = New RegExp
        'myRegExp.IgnoreCase = True
        'myRegExp.Global = True
        'myRegExp.Pattern = "regex"
        'myMatches = myRegExp.Execute(subjectString)
        'For Each myMatch In myMatches
        '    MsgBox(myMatch.Value)
        'Next
        Dim regex As Regex = New Regex("<table class=""listing"">.*?</table>")
        Dim regexStr As String = "<table class=""listing"">.*?<\/table>"
        Dim match As Match = regex.Match(str, regexStr, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        If match.Success Then
            'RichTextBox1.Text = match.Value
        Else
            MsgBox("Error in input" & match.Value)
        End If
        regex = New Regex("<a href="".*?</a>")
        regexStr = "<a href="".*?</a>"
        Dim matches As MatchCollection = regex.Matches(match.Value, regexStr, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        Dim i As Integer = 0
        Dim StrMatch As String = ""
        Dim MangaChapter() As String
        If matches.Count <> 0 Then
            ReDim MangaChapter(matches.Count)
            ReDim DownloadUrl(matches.Count)
        End If
        ListBox1.Items.Clear()
        While i < matches.Count
            MangaChapter(i) = matches(i).Value.Replace("<a href=", "")
            MangaChapter(i) = MangaChapter(i).Replace("""", "")
            regex = New Regex(">.*?</a>")
            regexStr = ">.*?</a>"
            match = regex.Match(MangaChapter(i), regexStr, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            StrMatch += MangaChapter(i)
            'ListBox1.Items.Add(MangaChapter(i).Split.First)
            DownloadUrl(i) = MangaChapter(i).Split.First
            ListBox1.Items.Add(match.Value.Replace("</a>", "").Replace(">", ""))
            i += 1
        End While
        'MsgBox(StrMatch)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim URL As String = "http://kissmanga.com/Manga/Haru-Koi"
        Dim request As HttpWebRequest = WebRequest.Create(URL)
        request.UserAgent = ".NET Framework Test Client"
        Dim response As HttpWebResponse = request.GetResponse()
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
        'Dim str As String = reader.ReadLine()
        Dim str As String = reader.ReadToEnd
        'Do While str.Length > 0
        '    Console.WriteLine(str)
        '    str += reader.ReadLine()
        'Loop
        'MsgBox(complete)
        'KissMangaImage("/Manga/Haru-Koi/Vol-001-Ch-004--Misa-s-Eatery?id=115540")
    End Sub

    Private Sub KissMangaImage(url As String)
        Dim KissMangaUrlCom As String = "http://kissmanga.com" + url
        Dim request As HttpWebRequest = WebRequest.Create(KissMangaUrlCom)
        request.UserAgent = ".NET Framework Test Client"
        Dim response As HttpWebResponse = request.GetResponse()
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
        Dim str As String = reader.ReadToEnd

        Dim regex As Regex = New Regex("lstImages.push.*?""\);")
        Dim regexStr As String = "lstImages.push.*?""\);"
        Dim match As MatchCollection = regex.Matches(str, regexStr, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        Dim i As Integer = 0
        Dim Client As New WebClient
        'MsgBox(match.Count)
        Dim Destination As String
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Path to Download Manga"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Destination = dialog.SelectedPath
        End If

        While i < match.Count
            Try
                Client.DownloadFile(match(i).Value.Replace("lstImages.push(""", "").Replace(""");", ""), Destination & i.ToString & ".jpg")
                Client.Dispose()
            Catch e As Exception
                MsgBox(e.Message)
            End Try
            'RichTextBox1.AppendText(match(i).Value.Replace("lstImages.push(""", "").Replace(""");", ""))
            i += 1
        End While
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        KissMangaImage(DownloadUrl(ListBox1.SelectedIndex))
    End Sub

End Class
