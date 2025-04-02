Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports Microsoft.Identity.Client
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class DownloadMail

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        ReceiveMails()
        Application.DoEvents()
    End Sub

#Region "class file "

    Class Cls_ALL_Mails
        Public Property value As Object
    End Class


    Class Cls_Mails
        Public Property receivedDateTime As String
        Public Property subject As String
    End Class
#End Region
    Private Sub ReceiveMails()

        Dim MessageId As String = ""
        Dim accessToken As String = ""
        Dim filename As String = ""

        Try
            Dim clientId As String = "ee69f21b-7806-41c8-b6bd-78996d926bfe"
            Dim clientSecret As String = "u2F8Q~PlIWcE4k5mlANrNCNXcmLrKg2Np27NMatB"
            Dim tenantId As String = "6f51e668-4272-4a4a-9738-369791e0bd08"


            Dim authority As String = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token"

            Dim app As ConfidentialClientApplication = ConfidentialClientApplicationBuilder.Create(clientId) _
                    .WithClientSecret(clientSecret) _
                    .WithAuthority(authority) _
                    .Build()

            Dim scopes() As String = {"https://graph.microsoft.com/.default"} ' Graph Scope for authentication
            Dim authResult As AuthenticationResult = app.AcquireTokenForClient(scopes).ExecuteAsync().Result
            accessToken = authResult.AccessToken

            Dim docDate As DateTime = DateTime.Now

            'Get All Mails with From Greater than equal to and less than filter 

            Dim strSkip As String = "&$top=400"

            Dim strFilter As String = "?$filter=isRead eq true and receivedDateTime ge " & docDate.AddDays(-1).ToString("yyyy-MM-dd") & "T00:00:00Z and receivedDateTime lt " & docDate.ToString("yyyy-MM-dd") & "T23:59:59Z and hasAttachments eq true&$select=sender & id" & strSkip
            Dim strURI As Uri = New Uri("https://graph.microsoft.com/v1.0/users/" & Trim(GlobalVar.mailId) & "/mailFolders/Inbox/messages" & strFilter)
            Dim wbrequest As HttpWebRequest = HttpWebRequest.Create(strURI)
            wbrequest.Method = "Get"
            wbrequest.Host = "graph.microsoft.com"
            wbrequest.Headers.Add("Authorization", accessToken)

            Dim wbresponse As HttpWebResponse
            wbresponse = CType(wbrequest.GetResponse(), HttpWebResponse)
            Dim responseFromServer As String = ""
            If wbresponse.StatusCode = 200 Or wbresponse.StatusCode = 201 Or wbresponse.StatusCode = 202 Then
                Using wbresponse
                    Dim dataStream As Stream = wbresponse.GetResponseStream()
                    Dim reader As StreamReader = New StreamReader(dataStream)
                    responseFromServer = reader.ReadToEnd()
                    reader.Close()
                    dataStream.Close()
                End Using

                wbresponse.Close()
            End If
            Dim data_all_Mails = JObject.FromObject(JsonConvert.DeserializeObject(responseFromServer)).ToObject(Of Cls_ALL_Mails)()

            If Not IsNothing(data_all_Mails) Then

                For i As Integer = 0 To data_all_Mails.value.count - 1
                    lbltotcnt.Text = i & "/" & data_all_Mails.value.count
                    Application.DoEvents()
                    If Not IsNothing(data_all_Mails.value(i)("id")) Then
                        MessageId = Trim(data_all_Mails.value(i)("id").ToString)
                        Dim sender_Address As String = Trim(data_all_Mails.value(i)("sender")("emailAddress")("address").ToString)
                        Dim sender_Name As String = Trim(data_all_Mails.value(i)("sender")("emailAddress")("name").ToString)
                        Dim strFileURI As Uri = New Uri("https://graph.microsoft.com/v1.0/users/" & Trim(GlobalVar.mailId) & "/mailFolders/Inbox/messages/" & MessageId & "/")
                        Dim wbflreq As HttpWebRequest = HttpWebRequest.Create(strFileURI)
                        wbflreq.Method = "Get"
                        wbflreq.Host = "graph.microsoft.com"
                        wbflreq.Headers.Add("Authorization", accessToken)

                        Dim wbflresp As HttpWebResponse
                        wbflresp = CType(wbflreq.GetResponse(), HttpWebResponse)
                        Dim flResponse As String = ""
                        If wbflresp.StatusCode = 200 Or wbflresp.StatusCode = 201 Or wbflresp.StatusCode = 202 Then
                            Using wbflresp
                                Dim dataStream As Stream = wbflresp.GetResponseStream()
                                Dim reader As StreamReader = New StreamReader(dataStream)
                                flResponse = reader.ReadToEnd()
                                reader.Close()
                                dataStream.Close()
                            End Using
                            wbflresp.Close()
                        End If
                        Dim date_File = JObject.FromObject(JsonConvert.DeserializeObject(flResponse)).ToObject(Of Cls_Mails)()
                        Dim receiveDate = date_File.receivedDateTime
                        Dim Subject = date_File.subject
                        lblsub.Text = Subject & "/" & receiveDate
                        Dwnmail(accessToken, MessageId, sender_Address & "_" & receiveDate)
                        Application.DoEvents()
                        If GlobalVar.FwdIDaddress <> "" And GlobalVar.FwdIDname <> "" Then
                            fwrdMail(accessToken, MessageId)
                        End If
                    End If

                Next
            End If

        Catch ex As Exception
            MsgBox(Err.Number & vbCrLf & Err.Description & vbCrLf & "Error on Msg Id :" & MessageId, MsgBoxStyle.Critical, " Auto Mail Downloading")

        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        If MsgBox("Are you sure to exit application (Y/N) ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        End
    End Sub

    Private Sub Btnchange_Click(sender As Object, e As EventArgs) Handles Btnchange.Click
        If Btnchange.Tag = 0 Then
            pnlpopup.Visible = True
            Btnchange.Tag = 1
            txtxmail.Text = GlobalVar.mailId
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        saverec()
        MsgBox("Please restart application.!", MsgBoxStyle.Critical, "")
        Application.Exit()
    End Sub
    Private Sub saverec()
        Dim strmsg As String = ""
        Dim Filepath As String = Application.StartupPath & "\mail.ini"
        Try
            pnlpopup.Visible = False
            Btnchange.Tag = 0
            If Trim(txtxmail.Text) = "" Then strmsg = "Invaild Mail Id" & vbCrLf

            If strmsg <> "" Then MsgBox(strmsg, MsgBoxStyle.Critical) : Exit Sub
            If System.IO.File.Exists(Filepath) = False Then
                System.IO.File.Create(Filepath).Dispose()
            End If

            strmsg = ""
            strmsg += Trim(txtxmail.Text) & vbCrLf
            System.IO.File.WriteAllText(Filepath, strmsg)

        Catch ex As Exception
            MsgBox(Err.Number & vbCrLf & Err.Description & MsgBoxStyle.Critical, " Auto Mail Downloading")
        End Try
    End Sub


    Private Sub DownloadMail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim filePath As String = Application.StartupPath & "/Eml Files/"
            If Not Directory.Exists(filePath) Then Directory.CreateDirectory(filePath)
            Dim frwdpath As String = Application.StartupPath & "/Forwardmail.ini"
            If System.IO.File.Exists(frwdpath) Then
                Dim strMsg As String = System.IO.File.ReadAllText(frwdpath)
                Dim strRslt As String() = Trim(strMsg).Replace(vbCrLf, "").Split(New String() {";"}, StringSplitOptions.RemoveEmptyEntries)
                GlobalVar.FwdIDname = Trim(strRslt(0))
                GlobalVar.FwdIDaddress = Trim(strRslt(1))
            End If
            openmail()
            lblmail.Text = GlobalVar.mailId
            ctrlCornerBorder(pnlpopup, 20)
        Catch ex As Exception
            MsgBox(Err.Number & vbCrLf & Err.Description & MsgBoxStyle.Critical, " Auto Mail Downloading")
        End Try

    End Sub
    Private Sub openmail()
        Dim Filepath As String = Application.StartupPath & "\mail.ini"
        Try
            If System.IO.File.Exists(Filepath) Then
                Dim strMsg As String = System.IO.File.ReadAllText(Filepath)

                Dim strRslt As String() = Trim(strMsg).Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                GlobalVar.mailId = Trim(strRslt(0))
            Else
                pnlpopup.Visible = True
                Btnchange.Tag = 1

            End If
        Catch ex As Exception
            MsgBox(Err.Number & vbCrLf & Err.Description & MsgBoxStyle.Critical, " Auto Mail Downloading")
        End Try

    End Sub


#Region "Download mail"
    Async Function Dwnmail(ByVal accessToken As String, ByVal messageId As String, ByVal sender As String) As Task
        sender = sender.Replace(":", "").Replace("*", "").Replace(",", "").Replace("/", "")
        Dim filePath As String = Application.StartupPath & "/Eml Files/"
        If Not Directory.Exists(filePath) Then Directory.CreateDirectory(filePath)
        Dim client As New HttpClient()
        client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/octet-stream"))

        Dim requestUrl As String = $"https://graph.microsoft.com/v1.0/users/" & Trim(GlobalVar.mailId) & "/messages/" & messageId & "/$value"

        Try
            Dim response As HttpResponseMessage = Await client.GetAsync(requestUrl)
            response.EnsureSuccessStatusCode()

            Dim emailContent As Byte() = Await response.Content.ReadAsByteArrayAsync()

            filePath = Application.StartupPath & "/Eml Files/" & sender & ".eml"
            System.IO.File.WriteAllBytes(filePath, emailContent)

            'MsgBox("Email saved to " & filePath)
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Function

    Private Sub btnclose_Click_1(sender As Object, e As EventArgs) Handles btnclose.Click
        If MsgBox("Are you sure to close (Y/N) ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then Exit Sub
        pnlpopup.Visible = False
        Btnchange.Tag = 0
    End Sub

    Private Sub btnemlfile_Click(sender As Object, e As EventArgs) Handles btnemlfile.Click
        Dim Folderpath As String = Application.StartupPath & "/Eml Files/"
        Process.Start(Folderpath)
    End Sub



#End Region


#Region "Forward  and Deleted Mail"

    Public Class EmailAddress
        Public Property name As String
        Public Property address As String
    End Class
    Public Class Root
        Public Property comment As String
        Public Property toRecipients As List(Of ToRecipient)
    End Class
    Public Class ToRecipient
        Public Property emailAddress As EmailAddress
    End Class
    Public Shared Function fwrdMail(ByVal accessToken As String, ByVal messageId As String)

        Dim obj As Root = New Root
        Dim objEml As EmailAddress = New EmailAddress
        Dim lstTo As List(Of ToRecipient) = New List(Of ToRecipient)
        Dim objTo As ToRecipient = New ToRecipient

        Try


            obj.comment = ""
            objEml.name = GlobalVar.FwdIDname
            objEml.address = GlobalVar.FwdIDaddress
            objTo.emailAddress = objEml
            lstTo.Add(objTo)
            obj.toRecipients = lstTo
            Dim strM = JsonConvert.SerializeObject(obj, Formatting.Indented)

            Dim url As String = $"Https://graph.microsoft.com/v1.0/users/" & Trim(GlobalVar.mailId) & "/mailFolders/Inbox/messages/" & messageId & "/forward"
            Using client1 As New HttpClient()
                client1.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                Dim requestBody As String = strM
                Dim content As New StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                Dim patchMethod As HttpMethod = New HttpMethod("POST")
                Dim request As New HttpRequestMessage(patchMethod, url) With {
                .Content = content
                }
                Dim frResponse As HttpResponseMessage = client1.SendAsync(request).Result
                If frResponse.IsSuccessStatusCode Then
                    'MsgBox("Message Forward Successfully.")
                    'Delete Mail
                    Dim DelURL As String = $"Https://graph.microsoft.com/v1.0/users/" & Trim(GlobalVar.mailId) & "/mailFolders/Inbox/messages/" & messageId
                    client1.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                    Dim Method As HttpMethod = New HttpMethod("DELETE")
                    Dim req As New HttpRequestMessage(Method, DelURL)
                    Dim DelRes As HttpResponseMessage = client1.SendAsync(req).Result
                    If DelRes.IsSuccessStatusCode Then
                        'MsgBox("Message Deleted.")
                    Else
                        MsgBox($"Failed to Deleted Message. Status code: {frResponse.StatusCode}")
                    End If
                Else
                    MsgBox($"Failed to Forward Message. Status code: {frResponse.StatusCode}")
                End If

            End Using

        Catch ex As Exception
            MsgBox(Err.Number & vbCrLf & Err.Description & MsgBoxStyle.Critical, " Auto Mail Downloading")
        End Try
    End Function


#End Region

End Class













