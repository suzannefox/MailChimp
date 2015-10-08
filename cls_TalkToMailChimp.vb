Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Collections.ObjectModel
Imports System.Web.Script.Serialization

Public Class cls_TalkToMailChimp

    Sub New(ByVal MyAPIKEY As String)
        Me.APIKEY = MyAPIKEY
    End Sub

#Region "Module variables"
    Private ErrorStatus As String = ""

    Private _apikey As String
    Public Property APIKEY() As String
        Get
            Return _apikey
        End Get
        Set(ByVal value As String)
            _apikey = value
            Dim Server = _apikey.Substring(_apikey.IndexOf("-") + 1)
            Me.URL = String.Format("https://{0}.api.mailchimp.com/2.0", Server)
        End Set
    End Property

    Private _url As String
    Public Property URL() As String
        Get
            Return _url
        End Get
        Set(ByVal value As String)
            _url = value
        End Set
    End Property

    ' -- Json text returned from MailChimp
    Private _JsonReturnText As String
    Public Property JsonReturnText() As String
        Get
            Return _JsonReturnText
        End Get

        Private Set(ByVal value As String)
            _JsonReturnText = value
            If Not _JsonReturnText.StartsWith("[") Then _JsonReturnText = "[" + _JsonReturnText
            If Not _JsonReturnText.EndsWith("]") Then _JsonReturnText += "]"
        End Set
    End Property

    ' -- Json text object for posting 
    Private _JsonPostText As String
    Public Property JsonPostText() As String
        Get
            Return _JsonPostText
        End Get
        Private Set(ByVal value As String)
            _JsonPostText = value
            _JsonPostText = _JsonPostText.Replace("[", "{")
            _JsonPostText = _JsonPostText.Replace("]", "}")
            _JsonPostText = _JsonPostText.Replace("~", Chr(34))
        End Set
    End Property

    ' -- DataType required from MailChimp
    Private _DataType As String
    Public Property DataType() As String
        Get
            Return _DataType
        End Get
        Private Set(ByVal value As String)
            _DataType = value.ToUpper
        End Set
    End Property

    Private _api_url As String
    Public Property api_url() As String
        Get
            Return _api_url
        End Get
        Private Set(ByVal value As String)
            _api_url = URL + value
        End Set
    End Property

    Private _MailChimpData As DataTable
    Public Property MailChimpData() As DataTable
        Get
            Return _MailChimpData
        End Get
        Private Set(ByVal value As DataTable)
            _MailChimpData = value
        End Set
    End Property

    Private _CampaignId As String
    Public Property Campaign_Id() As String
        Get
            Return _CampaignId
        End Get
        Set(ByVal value As String)
            _CampaignId = value
        End Set
    End Property

    Private _CampaignTitle As String
    Private Property Campaign_Title() As String
        Get
            Return _CampaignTitle
        End Get
        Set(ByVal value As String)
            _CampaignTitle = value
        End Set
    End Property
#End Region

#Region "Functions"

    ' -- Get Chimp Chatter Data, writes to JsonReturnText, MailChimpData
    Public Function Get_ChimpChatter() As String
        ErrorStatus = "OK"
        Dim ReturnVal As String = ""

        Try
            ' -- Get the JSON String
            Me.DataType = "chatter"
            Me.Build_JsonPostObject()
            ReturnVal = SendRequest(Me.api_url, Encoding.UTF8.GetBytes(Me.JsonPostText), "application/json", "POST")

            ' -- Deserialise to Datatable
            Dim dt As New DataTable

            dt.Columns.Add("Update_Time")
            dt.Columns.Add("Type")
            dt.Columns.Add("Message")

            Dim jss As New JavaScriptSerializer()
            Dim Chatter = jss.Deserialize(Of List(Of MC_ChimpChatter))(Me.JsonReturnText)

            For Each Chat In Chatter
                Dim dr As DataRow = dt.Rows.Add
                dr.Item("Update_Time") = Chat.update_time
                dr.Item("Type") = Chat.type
                dr.Item("Message") = Chat.message
            Next
            Me.MailChimpData = dt

        Catch ex As Exception
            ErrorStatus = "Get_ChimpChatter : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"

    End Function

    ' -- The public can only see the vanilla version of this function
    ' -- Get a list Of campaigns, Or optionally just the current campaigns details
    Public Function Get_Campaigns_List() As String
        Return Get_Campaigns_List(False, True)
    End Function

    ' -- Get a list of campaigns, or optionally just the current campaigns details
    Private Function Get_Campaigns_List(Optional CurrentCampaignOnly As Boolean = False,
                                        Optional GetBounceCount As Boolean = True) As String
        ErrorStatus = "OK"
        Dim ReturnVal As String = ""

        Try
            ' -- Get the JSON String
            If CurrentCampaignOnly = False Then
                Me.DataType = "campaign_list"
            Else
                Me.DataType = "campaign_title"
            End If

            Me.Build_JsonPostObject()
            ReturnVal = SendRequest(Me.api_url, Encoding.UTF8.GetBytes(Me.JsonPostText), "application/json", "POST")

            Dim dt As New DataTable

            dt.Columns.Add("Id")
            dt.Columns.Add("Web_Id")
            dt.Columns.Add("Title")
            dt.Columns.Add("Emails Sent")
            dt.Columns.Add("Send Time")
            dt.Columns.Add("From Name")
            dt.Columns.Add("Hard Bounces")
            dt.Columns.Add("Soft Bounces")
            dt.Columns.Add("Unsubscribes")
            dt.Columns.Add("Abuse Reports")

            Dim jss As New JavaScriptSerializer()
            Dim Campaigns = jss.Deserialize(Of List(Of MC_Campaigns_List_Topline))(Me.JsonReturnText)

            For Each Campaign In Campaigns

                For Each CampaignItem In Campaign.data
                    Dim dr As DataRow = dt.Rows.Add

                    dr.Item("Id") = CampaignItem.id
                    dr.Item("Web_Id") = CampaignItem.web_id
                    dr.Item("Title") = CampaignItem.title
                    dr.Item("Emails Sent") = CampaignItem.emails_sent
                    dr.Item("Send Time") = CampaignItem.send_time
                    dr.Item("From Name") = CampaignItem.from_name

                    ' -- If sent in the last 30 days then get stats
                    If GetBounceCount = True Then
                        If Not CampaignItem.send_time Is Nothing Then
                            If (CDate(CampaignItem.send_time) > DateTime.Today.AddDays(-30)) And CampaignItem.emails_sent > 50 Then
                                Dim CampaignDetails As New cls_TalkToMailChimp(Me.APIKEY)
                                CampaignDetails.Get_Campaign_Summary(CampaignItem.id)
                                dr.Item("Hard Bounces") = CampaignDetails.MailChimpData(0).Item("Hard_Bounces")
                                dr.Item("Soft Bounces") = CampaignDetails.MailChimpData(0).Item("Soft_Bounces")
                                dr.Item("Unsubscribes") = CampaignDetails.MailChimpData(0).Item("Unsubscribes")
                                dr.Item("Abuse Reports") = CampaignDetails.MailChimpData(0).Item("Abuse Reports")
                            End If
                        End If
                    End If
                Next
            Next

            Me.MailChimpData = dt
        Catch ex As Exception
            ErrorStatus = "Get_ChimpChatter : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"

    End Function

    ' -- For a specific campaign get summary stats
    Public Function Get_Campaign_Summary(ByVal Campaign_Id As String) As String
        ErrorStatus = "OK"
        Dim ReturnVal As String = ""

        Try
            ' -- Get the JSON String
            Me.DataType = "campaign_Summary"
            Me.Campaign_Id = Campaign_Id

            Me.Build_JsonPostObject()
            ReturnVal = SendRequest(Me.api_url, Encoding.UTF8.GetBytes(Me.JsonPostText), "application/json", "POST")

            Dim dt As New DataTable

            dt.Columns.Add("Hard_Bounces")
            dt.Columns.Add("Soft_Bounces")
            dt.Columns.Add("Unsubscribes")
            dt.Columns.Add("Abuse Reports")

            Dim jss As New JavaScriptSerializer()
            Dim CampaignStats = jss.Deserialize(Of List(Of MC_Campaign_Summary_Topline))(Me.JsonReturnText)

            Dim dr As DataRow = dt.Rows.Add

            dr.Item("Hard_Bounces") = CampaignStats(0).hard_bounces
            dr.Item("Soft_Bounces") = CampaignStats(0).soft_bounces
            dr.Item("Unsubscribes") = CampaignStats(0).unsubscribes
            dr.Item("Abuse Reports") = CampaignStats(0).abuse_reports

            Me.MailChimpData = dt
        Catch ex As Exception
            ErrorStatus = "Get_ChimpChatter : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"

    End Function

    ' -- For a specific campaign get hard bounces
    Public Function Get_Campaign_Bounces(Campaign_Id) As String
        ErrorStatus = "OK"
        Dim ReturnVal As String = ""

        Try
            ' -- Get the JSON String
            Me.DataType = "HARD_BOUNCES"
            Me.Campaign_Id = Campaign_Id
            If Me.Campaign_Id Is Nothing Then Me.Campaign_Id = "804cbcaaaa"

            ' -- Get campaign title -----------------------------------------
            Dim TempMC As New cls_TalkToMailChimp(Me.APIKEY)
            TempMC.Campaign_Id = Me.Campaign_Id
            TempMC.Get_Campaigns_List(True, False)
            Me.Campaign_Title = TempMC.MailChimpData.Rows(0).Item("Title")
            ' ---------------------------------------------------------------

            Dim dt As New DataTable
            dt.Columns.Add("Campaign")
            dt.Columns.Add("EMail")
            dt.Columns.Add("Status")

            ' -- Get the hard bounces
            Dim PageNo As Int16 = 0
            Dim BounceType As String = "hard"

            Do
                ' Create the JSON Object to POST
                Me.Build_JsonPostObject(PageNo, BounceType)
                ReturnVal = SendRequest(Me.api_url, Encoding.UTF8.GetBytes(Me.JsonPostText), "application/json", "POST")

                Dim TempTable As New DataTable
                TempTable = dt.Clone

                Dim jss As New JavaScriptSerializer()
                Dim Bounces = jss.Deserialize(Of List(Of MC_Bounce_Messages))(Me.JsonReturnText)

                For Each Bounce In Bounces
                    For Each BounceItem In Bounce.data
                        Dim dr As DataRow = TempTable.Rows.Add
                        dr.Item("Campaign") = Me.Campaign_Title
                        dr.Item("EMail") = BounceItem.member.Email
                        dr.Item("Status") = BounceItem.Status
                    Next
                Next

                ' -- When there are < 100 bounces then this is the end of the story
                dt.Merge(TempTable)
                If TempTable.Rows.Count < 100 Then Exit Do
                PageNo += 1

            Loop Until PageNo > 5 ' Just in case it goes wrong exit after 5 pages

            Me.MailChimpData = dt
        Catch ex As Exception
            ErrorStatus = "Get_Campaign_Bounces : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"
    End Function

    ' -- For a specific campaign get unsubscribers
    Public Function Get_Campaign_Unsubscribers(Campaign_Id) As String
        ErrorStatus = "OK"
        Dim ReturnVal As String = ""

        Try
            ' -- Get the JSON String
            Me.DataType = "UNSUBSCRIBERS"
            Me.Campaign_Id = Campaign_Id
            If Me.Campaign_Id Is Nothing Then Me.Campaign_Id = "804cbcaaaa"

            ' -- Get campaign title -----------------------------------------
            Dim TempMC As New cls_TalkToMailChimp(Me.APIKEY)
            TempMC.Campaign_Id = Me.Campaign_Id
            TempMC.Get_Campaigns_List(True, False)
            Me.Campaign_Title = TempMC.MailChimpData.Rows(0).Item("Title")
            ' ---------------------------------------------------------------

            Dim dt As New DataTable
            dt.Columns.Add("Campaign")
            dt.Columns.Add("EMail")
            dt.Columns.Add("Status")
            dt.Columns.Add("Reason")
            dt.Columns.Add("Reason_Text")

            ' -- Get the hard bounces
            Dim PageNo As Int16 = 0

            Do
                ' Create the JSON Object to POST
                Me.Build_JsonPostObject(PageNo)
                ReturnVal = SendRequest(Me.api_url, Encoding.UTF8.GetBytes(Me.JsonPostText), "application/json", "POST")

                Dim TempTable As New DataTable
                TempTable = dt.Clone

                Dim jss As New JavaScriptSerializer()
                Dim Bounces = jss.Deserialize(Of List(Of MC_Unsubscribe_Messages))(Me.JsonReturnText)

                For Each Bounce In Bounces
                    For Each BounceItem In Bounce.data
                        Dim dr As DataRow = TempTable.Rows.Add
                        dr.Item("Campaign") = Me.Campaign_Title
                        dr.Item("EMail") = BounceItem.member.Email
                        dr.Item("Status") = "unsubscriber"
                        dr.Item("Reason") = BounceItem.reason
                        dr.Item("Reason_Text") = BounceItem.reason_text
                    Next
                Next

                ' -- When there are < 100 bounces then this is the end of the story
                dt.Merge(TempTable)
                If TempTable.Rows.Count < 100 Then Exit Do
                PageNo += 1

            Loop Until PageNo > 5 ' Just in case it goes wrong exit after 5 pages

            Me.MailChimpData = dt
        Catch ex As Exception
            ErrorStatus = "Get_Campaign_Unsubscribers : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"
    End Function

    ' -- POST the Json Object as get the returned data
    Private Function SendRequest(ByVal uri As String, ByVal jsonDataBytes As Byte(), ByVal contentType As String, ByVal method As String) As String
        ErrorStatus = "OK"
        Try
            Me.JsonReturnText = ""
            Me.MailChimpData = Nothing

            Dim req As WebRequest = WebRequest.Create(uri)
            req.ContentType = contentType
            req.Method = method
            req.ContentLength = jsonDataBytes.Length


            Dim stream = req.GetRequestStream()
            stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
            stream.Close()

            Dim response = req.GetResponse().GetResponseStream()

            Dim reader As New StreamReader(response)
            Dim returnString = reader.ReadToEnd()
            reader.Close()
            response.Close()

            Me.JsonReturnText = returnString
        Catch ex As Exception
            ErrorStatus = "SendRequest : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"

    End Function

    ' -- Build the JSON Object to POST
    Private Function Build_JsonPostObject(Optional PageNo As Integer = Nothing, Optional BounceType As String = Nothing) As String
        ErrorStatus = "OK"

        Try
            Select Case Me.DataType
                Case "CHATTER"
                    Me.api_url = "/helper/chimp-chatter"                             ' api to call
                    Me.JsonPostText = String.Format("[~apikey~: ~{0}~]", APIKEY)     ' json object

                Case "CAMPAIGN_LIST"
                    Me.api_url = "/campaigns/list"                                   ' api to call
                    Me.JsonPostText = String.Format("[~apikey~: ~{0}~]", APIKEY)     ' json object

                Case "CAMPAIGN_SUMMARY"
                    Me.api_url = "/reports/summary"                                   ' api to call
                    Me.JsonPostText = String.Format("[~apikey~: ~{0}~,~cid~: ~{1}~]",
                                                    APIKEY, Me.Campaign_Id)           ' json object
                Case "CAMPAIGN_TITLE"
                    Me.api_url = "/campaigns/list"                                   ' api to call
                    Me.JsonPostText = String.Format("[~apikey~: ~{0}~,~filters~:[~campaign_id~: ~{1}~]]",
                                                    APIKEY, Me.Campaign_Id)           ' json object
                Case "HARD_BOUNCES"
                    Me.api_url = "/reports/sent-to"
                    Me.JsonPostText = String.Format("[~apikey~: ~{0}~,~cid~:~{1}~,~opts~: [~limit~: ~100~, ~start~: ~{2}~, ~status~: ~{3}~]]",
                                             APIKEY, Campaign_Id, PageNo, BounceType) ' json object
                Case "UNSUBSCRIBERS"
                    Me.api_url = "/reports/unsubscribes"
                    Me.JsonPostText = String.Format("[~apikey~: ~{0}~,~cid~:~{1}~,~opts~: [~limit~: ~100~, ~start~: ~{2}~]]",
                                             APIKEY, Campaign_Id, PageNo) ' json object

                Case Else
            End Select

        Catch ex As Exception
            ErrorStatus = "Build_JsonPostObject : " + ex.Message
            Return ex.Message
        End Try
        Return "OK"

    End Function
#End Region

#Region "MailChimp Classes for deserializing the JSON"

#Region "Campaign Summary"
    Private Class MC_Campaign_Summary_Topline
        Public Property syntax_errors As Integer
        Public Property hard_bounces As Integer
        Public Property soft_bounces As Integer
        Public Property unsubscribes As Integer
        Public Property abuse_reports As Integer
        Public Property forwards As Integer
        Public Property forwards_opens As Integer
        Public Property opens As Integer
        Public Property last_open As String
        Public Property unique_opens As Integer
        Public Property clicks As Integer
        Public Property unique_clicks As Integer
        Public Property last_click As String
        Public Property users_who_clicked As Integer
        Public Property emails_sent As Integer
        Public Property unique_likes As Integer
        Public Property recipient_likes As Integer
        Public Property facebook_likes As Integer
    End Class
#End Region
#Region "Campaign List"
    Private Class MC_Campaigns_List_Topline
        Public Property total As Integer
        Public Property data As MC_Campaigns_List_Topline_Datum()
    End Class

    Private Class MC_Campaigns_List_Topline_Datum
        Public Property id As String
        Public Property web_id As Integer
        Public Property list_id As String
        Public Property folder_id As Integer
        Public Property template_id As Integer
        Public Property content_type As String
        Public Property title As String
        Public Property type As String
        Public Property create_time As String
        Public Property send_time As String
        Public Property emails_sent As Integer
        Public Property status As String
        Public Property from_name As String
        Public Property from_email As String
        Public Property subject As String
        Public Property to_name As String
    End Class
#End Region

#Region "Unsubscriber-messages"

    Private Class MC_Unsubscribe_Messages
        Public Property total As Integer
        Public Property data As MC_Unsubscribe_Datum()
    End Class

    Private Class MC_Unsubscribe_Datum
        Public Property member As MC_Member
        Public Property reason As String
        Public Property reason_text As String
    End Class

#End Region

#Region "Bounce-messages"

    Private Class MC_Bounce_Messages
        Public Property total As Integer
        Public Property data As MC_Bounce_Datum()
    End Class

    Private Class MC_Bounce_Datum
        Public Property member As MC_Member
        Public Property Status As String
    End Class

    Private Class MC_Member
        Public Property Email As String
    End Class

#End Region
#Region "Chimp Chatter"

    Private Class MC_ChimpChatter
        Public Property message As String
        Public Property type As String
        Public Property url As String
        Public Property list_id As String
        Public Property campaign_id As String
        Public Property update_time As String
    End Class
End Class

#End Region
#End Region