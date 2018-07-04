Attribute VB_Name = "dsLaunchers"
Option Explicit


Sub setDatasourceVariables()

    If dataSource = "" Then
        If ActiveSheet.Name = "AdWords" Then
            dataSource = "AW"
        ElseIf ActiveSheet.Name = "Analytics" Then
            dataSource = "GA"
        ElseIf ActiveSheet.Name = "BingAds" Then
            dataSource = "AC"
        ElseIf ActiveSheet.Name = "Facebook" Then
            dataSource = "FB"
        ElseIf ActiveSheet.Name = "YouTube" Then
            dataSource = "YT"
        ElseIf ActiveSheet.Name = "Webmaster" Then
            dataSource = "GW"
        ElseIf ActiveSheet.Name = "Stripe" Then
            dataSource = "ST"
        ElseIf ActiveSheet.Name = "FacebookAds" Then
            dataSource = "FA"
        ElseIf ActiveSheet.Name = "MailChimp" Then
            dataSource = "MC"
        ElseIf ActiveSheet.Name = "TwitterAds" Then
            dataSource = "TA"
        End If
    End If

    If dataSource = "GA" Then
        Set configsheet = Analytics
        varsuffix = vbNullString
        dataSource = "GA"
        serviceName = "Google Analytics"
        appID = "GADG"
        moduleName = "Analytics Module"
        loginInfoCol = 1
        maxSimultaneousQueries = maxSimultaneousQueriesGA
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 2
        referToProfilesAs = "profiles"
        referToProfilesAsSing = "profile"
        referToAccountsAsSing = "account"
    ElseIf dataSource = "AW" Then
        Set configsheet = AdWords
        varsuffix = "AW"
        dataSource = "AW"
        serviceName = "AdWords"
        appID = "GADGAW"
        moduleName = "AdWords Module"
        loginInfoCol = 7
        maxSimultaneousQueries = maxSimultaneousQueriesAW
        Set varsSheetForDataSource = Sheets("varsAW")
        parameterColumnOffset = 3
        referToProfilesAs = "accounts"
        referToProfilesAsSing = "account"
        referToAccountsAsSing = "MCC"
    ElseIf dataSource = "AC" Then
        Set configsheet = BingAds
        varsuffix = "AC"
        dataSource = "AC"
        serviceName = "Bing Ads"
        appID = "GADGAC"
        moduleName = "Bing Ads Module"
        loginInfoCol = 13
        maxSimultaneousQueries = maxSimultaneousQueriesAC
        Set varsSheetForDataSource = Sheets("varsAC")
        parameterColumnOffset = 4
        referToProfilesAs = "accounts"
        referToProfilesAsSing = "account"
        referToAccountsAsSing = "MCC"
    ElseIf dataSource = "FB" Then
        Set configsheet = Facebook
        varsuffix = "FB"
        dataSource = "FB"
        serviceName = "Facebook"
        appID = "GADGFB"
        moduleName = "Facebook Module"
        loginInfoCol = 19
        maxSimultaneousQueries = maxSimultaneousQueriesFB
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 5
        referToProfilesAs = "pages/apps"
        referToProfilesAsSing = "page/app"
        referToAccountsAsSing = "account"
    ElseIf dataSource = "YT" Then
        Set configsheet = YouTube
        varsuffix = "YT"
        dataSource = "YT"
        serviceName = "YouTube"
        appID = "GADGYT"
        moduleName = "YouTube Module"
        loginInfoCol = 25
        maxSimultaneousQueries = maxSimultaneousQueriesYT
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 6
        referToProfilesAs = "videos/channels"
        referToProfilesAsSing = "video/channel"
        referToAccountsAsSing = "channel"
        'removed FL
    ElseIf dataSource = "TW" Then
        Set configsheet = Twitter
        varsuffix = "TW"
        dataSource = "TW"
        serviceName = "Twitter"
        appID = "GADGTW"
        moduleName = "Twitter Module"
        loginInfoCol = 37
        maxSimultaneousQueries = 1
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 8
        referToProfilesAs = ""
        referToProfilesAsSing = ""
        referToAccountsAsSing = ""
    ElseIf dataSource = "GW" Then
        Set configsheet = Webmaster
        varsuffix = "GW"
        dataSource = "GW"
        serviceName = "Search Console"
        appID = "GADGGW"
        moduleName = "Search Console Module"
        loginInfoCol = 43
        maxSimultaneousQueries = maxSimultaneousQueriesGW
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 9
        referToProfilesAs = "sites"
        referToProfilesAsSing = "site"
        referToAccountsAsSing = "sites"
    ElseIf dataSource = "ST" Then
        Set configsheet = Stripe
        varsuffix = "ST"
        dataSource = "ST"
        serviceName = "Stripe"
        appID = "GADGST"
        moduleName = "Stripe Module"
        loginInfoCol = 50
        maxSimultaneousQueries = maxSimultaneousQueriesST
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 10
        referToProfilesAs = "accounts"
        referToProfilesAsSing = "account"
        referToAccountsAsSing = "accounts"
    ElseIf dataSource = "FA" Then
        Set configsheet = FacebookAds
        varsuffix = "FA"
        dataSource = "FA"
        serviceName = "Facebook Ads"
        appID = "GADGFA"
        moduleName = "Facebook Ads Module"
        loginInfoCol = 57  ' +7
        maxSimultaneousQueries = 19
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 11  ' +1
        referToProfilesAs = "accounts"
        referToProfilesAsSing = "account"
        referToAccountsAsSing = "accounts"
    ElseIf dataSource = "MC" Then
        Set configsheet = MailChimp
        varsuffix = "MC"
        dataSource = "MC"
        serviceName = "MailChimp"
        appID = "GADGMC"
        moduleName = "MailChimp Module"
        loginInfoCol = 64  ' +7
        maxSimultaneousQueries = 19
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 12  ' +1
        referToProfilesAs = "campaigns/lists"
        referToProfilesAsSing = "campaign/list"
        referToAccountsAsSing = "accounts"
    ElseIf dataSource = "TA" Then
        Set configsheet = TwitterAds
        varsuffix = "TA"
        dataSource = "TA"
        serviceName = "Twitter Ads"
        appID = "GADGTA"
        moduleName = "Twitter Ads Module"
        loginInfoCol = 71  ' +7
        maxSimultaneousQueries = 19
        Set varsSheetForDataSource = Sheets("vars")
        parameterColumnOffset = 13  ' +1
        referToProfilesAs = "accounts"
        referToProfilesAsSing = "account"
        referToAccountsAsSing = "accounts"

    ElseIf debugMode = True Then
        MsgBox "Datasource not set!"
        '        Stop
    End If

End Sub

Sub logoClickGA()
    On Error Resume Next
    If Analytics.Visible = xlSheetVisible Then Analytics.Select
End Sub
Sub logoClickAW()
    On Error Resume Next
    If AdWords.Visible = xlSheetVisible Then AdWords.Select
End Sub
Sub logoClickAC()
    On Error Resume Next
    If BingAds.Visible = xlSheetVisible Then BingAds.Select
End Sub
Sub logoClickFB()
    On Error Resume Next
    If Facebook.Visible = xlSheetVisible Then Facebook.Select
End Sub
Sub logoClickYT()
    On Error Resume Next
    If YouTube.Visible = xlSheetVisible Then YouTube.Select
End Sub
Sub logoClickTW()
    On Error Resume Next
    If Twitter.Visible = xlSheetVisible Then Twitter.Select
End Sub
Sub logoClickGW()
    On Error Resume Next
    If Webmaster.Visible = xlSheetVisible Then Webmaster.Select
End Sub
Sub logoClickST()
    On Error Resume Next
    If Stripe.Visible = xlSheetVisible Then Stripe.Select
End Sub
Sub logoClickFA()
    On Error Resume Next
    If Stripe.Visible = xlSheetVisible Then Stripe.Select
End Sub
Sub logoClickMC()
    On Error Resume Next
    If MailChimp.Visible = xlSheetVisible Then Stripe.Select
End Sub
Sub logoClickTA()
    On Error Resume Next
    If TwitterAds.Visible = xlSheetVisible Then Stripe.Select
End Sub


Sub selectStartDateGA()
    dataSource = "GA"
    Call selectStartDate
End Sub
Sub selectEndDateGA()
    dataSource = "GA"
    Call selectEndDate
End Sub
Sub selectStartDateAW()
    dataSource = "AW"
    Call selectStartDate
End Sub
Sub selectEndDateAW()
    dataSource = "AW"
    Call selectEndDate
End Sub
Sub selectStartDateAC()
    dataSource = "AC"
    Call selectStartDate
End Sub
Sub selectEndDateAC()
    dataSource = "AC"
    Call selectEndDate
End Sub
Sub selectStartDateFB()
    dataSource = "FB"
    Call selectStartDate
End Sub
Sub selectStartDateST()
    dataSource = "ST"
    Call selectStartDate
End Sub
Sub selectStartDateFA()
    dataSource = "FA"
    Call selectStartDate
End Sub
Sub selectStartDateMC()
    dataSource = "MC"
    Call selectStartDate
End Sub
Sub selectStartDateTA()
    dataSource = "TA"
    Call selectStartDate
End Sub



Sub selectEndDateFB()
    dataSource = "FB"
    Call selectEndDate
End Sub
Sub selectStartDateYT()
    dataSource = "YT"
    Call selectStartDate
End Sub
Sub selectEndDateYT()
    dataSource = "YT"
    Call selectEndDate
End Sub
Sub selectStartDateGW()
    dataSource = "GW"
    Call selectStartDate
End Sub
Sub selectEndDateGW()
    dataSource = "GW"
    Call selectEndDate
End Sub
Sub selectEndDateST()
    dataSource = "ST"
    Call selectEndDate
End Sub
Sub selectEndDateFA()
    dataSource = "FA"
    Call selectEndDate
End Sub
Sub selectEndDateMC()
    dataSource = "MC"
    Call selectEndDate
End Sub
Sub selectEndDateTA()
    dataSource = "TA"
    Call selectEndDate
End Sub





Sub dateRangeTypeChangeGA()
    dataSource = "GA"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeAW()
    dataSource = "AW"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeAC()
    dataSource = "AC"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeFB()
    dataSource = "FB"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeYT()
    dataSource = "YT"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeGW()
    dataSource = "GW"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeST()
    dataSource = "ST"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeFA()
    dataSource = "FA"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeMC()
    dataSource = "MC"
    Call dateRangeTypeChange
End Sub
Sub dateRangeTypeChangeTA()
    dataSource = "TA"
    Call dateRangeTypeChange
End Sub



Sub changeDateRangeTypeXGA()
    dataSource = "GA"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXAW()
    dataSource = "AW"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXAC()
    dataSource = "AC"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXFB()
    dataSource = "FB"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXYT()
    dataSource = "YT"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXGW()
    dataSource = "GW"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXST()
    dataSource = "ST"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXFA()
    dataSource = "FA"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXMC()
    dataSource = "MC"
    Call changeDateRangeTypeX
End Sub
Sub changeDateRangeTypeXTA()
    dataSource = "TA"
    Call changeDateRangeTypeX
End Sub





Sub updateVisibilitydrdga()
    Call updateVisibilityOfDropdowns("drdga")
    Call checkSelectedGAfields
End Sub
Sub updateVisibilitydrsdga()
    Call updateVisibilityOfDropdowns("drsdga")
    Call checkSelectedGAfields
End Sub
Sub updateVisibilitydrmga()
    Call updateVisibilityOfDropdowns("drmga")
    Call checkSelectedGAfields
End Sub
Sub updateVisibilitydrdaw()
    Call updateVisibilityOfDropdowns("drdaw")
    Call checkSelectedAWfields
End Sub
Sub updateVisibilitydrsdaw()
    Call updateVisibilityOfDropdowns("drsdaw")
    Call checkSelectedAWfields
End Sub
Sub updateVisibilitydrmaw()
    Call updateVisibilityOfDropdowns("drmaw")
    Call checkSelectedAWfields
End Sub
Sub updateVisibilitydrdac()
    Call updateVisibilityOfDropdowns("drdac")
    Call checkSelectedACfields
End Sub
Sub updateVisibilitydrsdac()
    Call updateVisibilityOfDropdowns("drsdac")
    Call checkSelectedACfields
End Sub
Sub updateVisibilitydrmac()
    Call updateVisibilityOfDropdowns("drmac")
    Call checkSelectedACfields
End Sub
Sub updateVisibilitydrdfb()
    Call updateVisibilityOfDropdowns("drdfb")
End Sub
Sub updateVisibilitydrsdfb()
    Call updateVisibilityOfDropdowns("drsdfb")
End Sub
Sub updateVisibilitydrmfb()
    Call updateVisibilityOfDropdowns("drmfb")
End Sub
Sub updateVisibilitydrdyt()
    Call updateVisibilityOfDropdowns("drdyt")
    Call checkSelectedYTfields
End Sub
Sub updateVisibilitydrsdyt()
    Call updateVisibilityOfDropdowns("drsdyt")
    Call checkSelectedYTfields
End Sub
Sub updateVisibilitydrmyt()
    Call updateVisibilityOfDropdowns("drmyt")
    Call checkSelectedYTfields
End Sub
Sub updateVisibilitydrdfl()
    Call updateVisibilityOfDropdowns("drdfl")
End Sub
Sub updateVisibilitydrsdfl()
    Call updateVisibilityOfDropdowns("drsdfl")
End Sub
Sub updateVisibilitydrmfl()
    Call updateVisibilityOfDropdowns("drmfl")
End Sub
Sub updateVisibilitydrdgw()
    Call updateVisibilityOfDropdowns("drdgw")
    Call checkSelectedGWfields
End Sub
Sub updateVisibilitydrsdgw()
    Call updateVisibilityOfDropdowns("drsdgw")
    Call checkSelectedGWfields
End Sub
Sub updateVisibilitydrmgw()
    Call updateVisibilityOfDropdowns("drmgw")
    Call checkSelectedGWfields
End Sub
Sub updateVisibilitydrdst()
    Call updateVisibilityOfDropdowns("drdst")
End Sub
Sub updateVisibilitydrsdst()
    Call updateVisibilityOfDropdowns("drsdst")
End Sub
Sub updateVisibilitydrmst()
    Call updateVisibilityOfDropdowns("drmst")
End Sub
Sub updateVisibilitydrdfa()
    Call updateVisibilityOfDropdowns("drdfa")
End Sub
Sub updateVisibilitydrsdfa()
    Call updateVisibilityOfDropdowns("drsdfa")
End Sub
Sub updateVisibilitydrmfa()
    Call updateVisibilityOfDropdowns("drmfa")
End Sub
Sub updateVisibilitydrdmc()
    Call updateVisibilityOfDropdowns("drdmc")
    Call checkSelectedMCfields
End Sub
Sub updateVisibilitydrsdmc()
    Call updateVisibilityOfDropdowns("drsdmc")
    Call checkSelectedMCfields
End Sub
Sub updateVisibilitydrmmc()
    Call updateVisibilityOfDropdowns("drmmc")
    Call checkSelectedMCfields
End Sub
Sub updateVisibilitydrdta()
    Call updateVisibilityOfDropdowns("drdta")
    Call checkSelectedTAfields
End Sub
Sub updateVisibilitydrsdta()
    Call updateVisibilityOfDropdowns("drsdta")
    Call checkSelectedTAfields
End Sub
Sub updateVisibilitydrmta()
    Call updateVisibilityOfDropdowns("drmta")
    Call checkSelectedTAfields
End Sub






Sub updateFieldSelectionsGA()
    dataSource = "GA"
    Call updateVisibilityOfDropdowns("drdga")
    Call updateVisibilityOfDropdowns("drsdga")
    Call updateVisibilityOfDropdowns("drmga")
End Sub

Sub updateFieldSelectionsAW()
    dataSource = "AW"
    Call updateVisibilityOfDropdowns("drdaw")
    Call updateVisibilityOfDropdowns("drsdaw")
    Call updateVisibilityOfDropdowns("drmaw")
End Sub
Sub updateFieldSelectionsAC()
    dataSource = "AC"
    Call updateVisibilityOfDropdowns("drdac")
    Call updateVisibilityOfDropdowns("drsdac")
    Call updateVisibilityOfDropdowns("drmac")
End Sub
Sub updateFieldSelectionsFB()
    dataSource = "FB"
    Call updateVisibilityOfDropdowns("drdfb")
    Call updateVisibilityOfDropdowns("drsdfb")
    Call updateVisibilityOfDropdowns("drmfb")
End Sub
Sub updateFieldSelectionsYT()
    dataSource = "YT"
    Call updateVisibilityOfDropdowns("drdyt")
    Call updateVisibilityOfDropdowns("drsdyt")
    Call updateVisibilityOfDropdowns("drmyt")
    Range("metricsCalcYT").Calculate
End Sub
Sub updateFieldSelectionsST()
    dataSource = "ST"
    Call updateVisibilityOfDropdowns("drdst")
    Call updateVisibilityOfDropdowns("drsdst")
    Call updateVisibilityOfDropdowns("drmst")
    Range("metricsCalcST").Calculate
End Sub
Sub updateFieldSelectionsFA()
    dataSource = "FA"
    Call updateVisibilityOfDropdowns("drdfa")
    Call updateVisibilityOfDropdowns("drsdfa")
    Call updateVisibilityOfDropdowns("drmfa")
    Range("metricsCalcFA").Calculate
End Sub
Sub updateFieldSelectionsMC()
    dataSource = "MC"
    Call updateVisibilityOfDropdowns("drdmc")
    Call updateVisibilityOfDropdowns("drsdmc")
    Call updateVisibilityOfDropdowns("drmmc")
    Range("metricsCalcMC").Calculate
End Sub
Sub updateFieldSelectionsTA()
    dataSource = "TA"
    Call updateVisibilityOfDropdowns("drdta")
    Call updateVisibilityOfDropdowns("drsdta")
    Call updateVisibilityOfDropdowns("drmta")
    Range("metricsCalcTA").Calculate
End Sub






Sub clearFieldSelectionsGA()
    dataSource = "GA"
    Call clearFieldSelections
    Analytics.Shapes("illegalFieldsWarning").Visible = False
    Analytics.Shapes("fieldsOKnote").Visible = False
End Sub
Sub clearFieldSelectionsAW()
    dataSource = "AW"
    Call clearFieldSelections
    Call unprotectSheets
    AdWords.Shapes("illegalFieldsWarningAW").Visible = False
    AdWords.Shapes("fieldsOKnote").Visible = False
    Call protectSheets
End Sub
Sub clearFieldSelectionsAC()
    On Error Resume Next
    dataSource = "AC"
    Call clearFieldSelections
    Call unprotectSheets
    BingAds.Shapes("illegalFieldsWarningAC").Visible = False
    BingAds.Shapes("fieldsOKnote").Visible = False
    Call protectSheets
End Sub
Sub clearFieldSelectionsFB()
    On Error Resume Next
    dataSource = "FB"
    Call clearFieldSelections
End Sub
Sub clearFieldSelectionsYT()
    On Error Resume Next
    dataSource = "YT"
    Call clearFieldSelections
    Call unprotectSheets
    YouTube.Shapes("illegalFieldsWarningYT").Visible = False
    YouTube.Shapes("fieldsOKnote").Visible = False
    Call protectSheets
End Sub
Sub clearFieldSelectionsGW()
    On Error Resume Next
    dataSource = "GW"
    Call clearFieldSelections
    Call unprotectSheets
    Webmaster.Shapes("illegalFieldsWarningGW").Visible = False
    Webmaster.Shapes("fieldsOKnote").Visible = False
    Call protectSheets
End Sub
Sub clearFieldSelectionsST()
    On Error Resume Next
    dataSource = "ST"
    Call clearFieldSelections
End Sub
Sub clearFieldSelectionsFA()
    On Error Resume Next
    dataSource = "FA"
    Call clearFieldSelections
End Sub
Sub clearFieldSelectionsMC()
    On Error Resume Next
    dataSource = "MC"
    Call clearFieldSelections
    Call unprotectSheets
    MailChimp.Shapes("illegalFieldsWarningMC").Visible = False
    MailChimp.Shapes("fieldsOKnote").Visible = False
    Call protectSheets
End Sub
Sub clearFieldSelectionsTA()
    On Error Resume Next
    dataSource = "TA"
    Call clearFieldSelections
    Call unprotectSheets
    TwitterAds.Shapes("illegalFieldsWarningTA").Visible = False
    TwitterAds.Shapes("fieldsOKnote").Visible = False
    Call protectSheets
End Sub


Sub clearProfileSelectionsGA()
    dataSource = "GA"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsAW()
    dataSource = "AW"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsAC()
    dataSource = "AC"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsFB()
    dataSource = "FB"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsYT()
    dataSource = "YT"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsGW()
    dataSource = "GW"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsST()
    dataSource = "ST"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsFA()
    dataSource = "FA"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsMC()
    dataSource = "MC"
    Call clearProfileSelections
End Sub
Sub clearProfileSelectionsTA()
    dataSource = "TA"
    Call clearProfileSelections
End Sub





Sub selectAllProfilesGA()
    dataSource = "GA"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesAW()
    dataSource = "AW"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesAC()
    dataSource = "AC"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesFB()
    dataSource = "FB"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesYT()
    dataSource = "YT"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesGW()
    dataSource = "GW"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesST()
    dataSource = "ST"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesFA()
    dataSource = "FA"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesMC()
    dataSource = "MC"
    Call selectAllProfiles
End Sub
Sub selectAllProfilesTA()
    dataSource = "TA"
    Call selectAllProfiles
End Sub



Sub updateProfileSelectionsGA()
    dataSource = "GA"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsAW()
    dataSource = "AW"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsAC()
    dataSource = "AC"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsFB()
    dataSource = "FB"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsYT()
    dataSource = "YT"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsGW()
    dataSource = "GW"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsST()
    dataSource = "ST"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsFA()
    dataSource = "FA"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsMC()
    dataSource = "MC"
    Call updateProfileSelections
    Call protectSheets
End Sub
Sub updateProfileSelectionsTA()
    dataSource = "TA"
    Call updateProfileSelections
    Call protectSheets
End Sub





Sub updateProfileSelectionsGAbutton()
    dataSource = "GA"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsAWbutton()
    dataSource = "AW"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsACbutton()
    dataSource = "AC"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsFBbutton()
    dataSource = "FB"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsYTbutton()
    dataSource = "YT"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsGWbutton()
    dataSource = "GW"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsSTbutton()
    dataSource = "ST"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsFAbutton()
    dataSource = "FA"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsMCbutton()
    dataSource = "MC"
    Call updateProfileSelectionsButton
End Sub
Sub updateProfileSelectionsTAbutton()
    dataSource = "TA"
    Call updateProfileSelectionsButton
End Sub





Sub OAuthLoginFB()
    dataSource = "FB"
    Call OAuthLogin
End Sub
Sub OAuthLoginGA()
    dataSource = "GA"
    Call OAuthLogin
End Sub
Sub OAuthLoginGA1()
    dataSource = "GA"
    Call OAuthLogin
End Sub
Sub OAuthLoginAW()
    dataSource = "AW"
    Call OAuthLogin
End Sub
Sub OAuthLoginYT()
    dataSource = "YT"
    Call OAuthLogin
End Sub
Sub OAuthLoginAC()
    dataSource = "AC"
    Call OAuthLogin
End Sub
Sub OAuthLoginTW()
    dataSource = "TW"
    Call OAuthLogin
End Sub
Sub OAuthLoginGW()
    dataSource = "GW"
    Call OAuthLogin
End Sub
Sub OAuthLoginST()
    dataSource = "ST"
    Call OAuthLogin
End Sub
Sub OAuthLoginFA()
    dataSource = "FA"
    Call OAuthLogin
End Sub
Sub OAuthLoginMC()
    dataSource = "MC"
    Call OAuthLogin
End Sub
Sub OAuthLoginTA()
    dataSource = "TA"
    Call OAuthLogin
End Sub




Sub addLoginGA()
    dataSource = "GA"
    Call addLogin
End Sub
Sub addLoginAW()
    dataSource = "AW"
    Call addLogin
End Sub
Sub addLoginAC()
    dataSource = "AC"
    Call addLogin
End Sub
Sub addLoginFB()
    dataSource = "FB"
    Call addLogin
End Sub
Sub addLoginYT()
    dataSource = "YT"
    Call addLogin
End Sub
Sub addLoginGW()
    dataSource = "GW"
    Call addLogin
End Sub
Sub addLoginST()
    dataSource = "ST"
    Call addLogin
End Sub
Sub addLoginFA()
    dataSource = "FA"
    Call addLogin
End Sub
Sub addLoginMC()
    dataSource = "MC"
    Call addLogin
End Sub
Sub addLoginTA()
    dataSource = "TA"
    Call addLogin
End Sub



Sub refreshProfileListGA()
    dataSource = "GA"
    Call refreshProfileList
End Sub
Sub refreshProfileListAW()
    dataSource = "AW"
    Call refreshProfileList
End Sub
Sub refreshProfileListAC()
    dataSource = "AC"
    Call refreshProfileList
End Sub
Sub refreshProfileListFB()
    dataSource = "FB"
    Call refreshProfileList
End Sub
Sub refreshProfileListYT()
    dataSource = "YT"
    Call refreshProfileList
End Sub
Sub refreshProfileListGW()
    dataSource = "GW"
    Call refreshProfileList
End Sub
Sub refreshProfileListST()
    dataSource = "ST"
    Call refreshProfileList
End Sub
Sub refreshProfileListFA()
    dataSource = "FA"
    Call refreshProfileList
End Sub
Sub refreshProfileListMC()
    dataSource = "MC"
    Call refreshProfileList
End Sub
Sub refreshProfileListTA()
    dataSource = "TA"
    Call refreshProfileList
End Sub


Sub launchLoginStatusBoxGA1()
    Call launchLoginStatusBoxGA
End Sub
Sub launchLoginStatusBoxGA()
    dataSource = "GA"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxAW1()
    Call launchLoginStatusBoxAW
End Sub
Sub launchLoginStatusBoxAW()
    dataSource = "AW"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxAC1()
    Call launchLoginStatusBoxAC
End Sub
Sub launchLoginStatusBoxAC()
    dataSource = "AC"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxFB1()
    Call launchLoginStatusBoxFB
End Sub
Sub launchLoginStatusBoxFB()
    dataSource = "FB"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxYT()
    dataSource = "YT"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxGW()
    dataSource = "GW"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxST()
    dataSource = "ST"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxST1()
    Call launchLoginStatusBoxST
End Sub
Sub launchLoginStatusBoxFA()
    dataSource = "FA"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxFA1()
    Call launchLoginStatusBoxFA
End Sub
Sub launchLoginStatusBoxMC()
    dataSource = "MC"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxMC1()
    Call launchLoginStatusBoxMC
End Sub
Sub launchLoginStatusBoxTA()
    dataSource = "TA"
    Call launchLoginStatusBox
End Sub
Sub launchLoginStatusBoxTA1()
    Call launchLoginStatusBoxTA
End Sub





Sub launchFilterUFGA()
    dataSource = "GA"
    Call launchFilterUF
End Sub
Sub launchFilterUFAW()
    dataSource = "AW"
    Call launchFilterUF
End Sub
Sub launchFilterUFAC()
    dataSource = "AC"
    Call launchFilterUF
End Sub
Sub launchFilterUFFB()
    dataSource = "FB"
    Call launchFilterUF
End Sub
Sub launchFilterUFYT()
    dataSource = "YT"
    Call launchFilterUF
End Sub
Sub launchFilterUFGW()
    dataSource = "GW"
    Call launchFilterUF
End Sub
Sub launchFilterUFST()
    dataSource = "ST"
    Call launchFilterUF
End Sub
Sub launchFilterUFFA()
    dataSource = "FA"
    Call launchFilterUF
End Sub
Sub launchFilterUFMC()
    dataSource = "MC"
    Call launchFilterUF
End Sub
Sub launchFilterUFTA()
    dataSource = "TA"
    Call launchFilterUF
End Sub



Sub clearFiltersGA()
    dataSource = "GA"
    Call clearFilters
End Sub
Sub clearFiltersAW()
    dataSource = "AW"
    Call clearFilters
End Sub
Sub clearFiltersAC()
    dataSource = "AC"
    Call clearFilters
End Sub
Sub clearFiltersFB()
    dataSource = "FB"
    Call clearFilters
End Sub
Sub clearFiltersYT()
    dataSource = "YT"
    Call clearFilters
End Sub
Sub clearFiltersGW()
    dataSource = "GW"
    Call clearFilters
End Sub
Sub clearFiltersST()
    dataSource = "ST"
    Call clearFilters
End Sub
Sub clearFiltersFA()
    dataSource = "FA"
    Call clearFilters
End Sub
Sub clearFiltersMC()
    dataSource = "MC"
    Call clearFilters
End Sub
Sub clearFiltersTA()
    dataSource = "TA"
    Call clearFilters
End Sub





Sub logoutGA()
    dataSource = "GA"
    Call logout(True)
End Sub
Sub logoutAW()
    dataSource = "AW"
    Call logout(True)
End Sub
Sub logoutAC()
    dataSource = "AC"
    Call logout(False)
End Sub
Sub logoutFB()
    dataSource = "FB"
    Call logout(True)
End Sub
Sub logoutYT()
    dataSource = "YT"
    Call logout(True)
End Sub
Sub logoutTW()
    dataSource = "TW"
    Call logout(True)
End Sub
Sub logoutGW()
    dataSource = "GW"
    Call logout(True)
End Sub
Sub logoutST()
    dataSource = "ST"
    Call logout(True)
End Sub
Sub logoutFA()
    dataSource = "FA"
    Call logout(True)
End Sub
Sub logoutMC()
    dataSource = "MC"
    Call logout(True)
End Sub
Sub logoutTA()
    dataSource = "TA"
    Call logout(True)
End Sub




Sub runReportGA()
    dataSource = "GA"
    Call runReport
End Sub
Sub runReportAW()
    dataSource = "AW"
    Call runReport
End Sub
Sub runReportAC()
    dataSource = "AC"
    Call runReport
End Sub
Sub runReportFB()
    dataSource = "FB"
    Call runReport
End Sub
Sub runReportYT()
    dataSource = "YT"
    Call runReport
End Sub
Sub runReportGW()
    dataSource = "GW"
    Call runReport
End Sub
Sub runReportST()
    dataSource = "ST"
    Call runReport
End Sub
Sub runReportFA()
    dataSource = "FA"
    Call runReport
End Sub
Sub runReportMC()
    dataSource = "MC"
    Call runReport
End Sub
Sub runReportTA()
    dataSource = "TA"
    Call runReport
End Sub



Sub aggregateQueryAW()
    dataSource = "AW"
    Call aggregateQuery
End Sub
Sub aggregateQueryGA()
    dataSource = "GA"
    Call aggregateQuery
End Sub
Sub aggregateQueryAC()
    dataSource = "AC"
    Call aggregateQuery
End Sub
Sub aggregateQueryFB()
    dataSource = "FB"
    Call aggregateQuery
End Sub
Sub aggregateQueryYT()
    dataSource = "YT"
    Call aggregateQuery
End Sub
Sub aggregateQueryGW()
    dataSource = "GW"
    Call aggregateQuery
End Sub
Sub aggregateQueryST()
    dataSource = "ST"
    Call aggregateQuery
End Sub
Sub aggregateQueryFA()
    dataSource = "FA"
    Call aggregateQuery
End Sub
Sub aggregateQueryMC()
    dataSource = "MC"
    Call aggregateQuery
End Sub
Sub aggregateQueryTA()
    dataSource = "TA"
    Call aggregateQuery
End Sub



Sub dimensionQueryAW()
    dataSource = "AW"
    Call dimensionQuery
End Sub
Sub dimensionQueryGA()
    dataSource = "GA"
    Call dimensionQuery
End Sub
Sub dimensionQueryAC()
    dataSource = "AC"
    Call dimensionQuery
End Sub
Sub dimensionQueryFB()
    dataSource = "FB"
    Call dimensionQuery
End Sub
Sub dimensionQueryYT()
    dataSource = "YT"
    Call dimensionQuery
End Sub
Sub dimensionQueryGW()
    dataSource = "GW"
    Call dimensionQuery
End Sub
Sub dimensionQueryST()
    dataSource = "ST"
    Call dimensionQuery
End Sub
Sub dimensionQueryFA()
    dataSource = "FA"
    Call dimensionQuery
End Sub
Sub dimensionQueryMC()
    dataSource = "MC"
    Call dimensionQuery
End Sub
Sub dimensionQueryTA()
    dataSource = "TA"
    Call dimensionQuery
End Sub



Sub checkSelectedACfields()
    dataSource = "AC"
    Call checkSelectedFields
End Sub
Sub checkSelectedAWfields()
    dataSource = "AW"
    Call checkSelectedFields
End Sub
Sub checkSelectedGAfields()
    dataSource = "GA"
    Call checkSelectedFields
End Sub
Sub checkSelectedGWfields()
    dataSource = "GW"
    Call checkSelectedFields
End Sub
Sub checkSelectedYTfields()
    dataSource = "YT"
    Call checkSelectedFields
End Sub
Sub checkSelectedMCfields()
    dataSource = "MC"
    Call checkSelectedFields
End Sub
Sub checkSelectedTAfields()
    dataSource = "TA"
    Call checkSelectedFields
End Sub




Sub checkSelectedACfieldsUI()
    dataSource = "AC"
    Call checkSelectedFields(True)
End Sub
Sub checkSelectedAWfieldsUI()
    dataSource = "AW"
    Call checkSelectedFields(True)
End Sub
Sub checkSelectedGAfieldsUI()
    dataSource = "GA"
    Call checkSelectedFields(True)
End Sub
Sub checkSelectedGWfieldsUI()
    dataSource = "GW"
    Call checkSelectedFields(True)
End Sub
Sub checkSelectedYTfieldsUI()
    dataSource = "YT"
    Call checkSelectedFields(True)
End Sub
Sub checkSelectedMCfieldsUI()
    dataSource = "MC"
    Call checkSelectedFields(True)
End Sub
Sub checkSelectedTAfieldsUI()
    dataSource = "TA"
    Call checkSelectedFields(True)
End Sub



Sub clearFieldSelections()
    On Error Resume Next
    Application.ScreenUpdating = False
    If dataSource = "GA" Then
        Call updateVisibilityOfDropdowns("drdga", True)
        Call updateVisibilityOfDropdowns("drsdga", True)
        Call updateVisibilityOfDropdowns("drmga", True)
    ElseIf dataSource = "AW" Then
        Call updateVisibilityOfDropdowns("drdaw", True)
        Call updateVisibilityOfDropdowns("drsdaw", True)
        Call updateVisibilityOfDropdowns("drmaw", True)
    ElseIf dataSource = "AC" Then
        Call updateVisibilityOfDropdowns("drdac", True)
        Call updateVisibilityOfDropdowns("drsdac", True)
        Call updateVisibilityOfDropdowns("drmac", True)
    ElseIf dataSource = "FB" Then
        Call updateVisibilityOfDropdowns("drdfb", True)
        Call updateVisibilityOfDropdowns("drsdfb", True)
        Call updateVisibilityOfDropdowns("drmfb", True)
    ElseIf dataSource = "YT" Then
        Call updateVisibilityOfDropdowns("drdyt", True)
        Call updateVisibilityOfDropdowns("drsdyt", True)
        Call updateVisibilityOfDropdowns("drmyt", True)
    ElseIf dataSource = "GW" Then
        Call updateVisibilityOfDropdowns("drdgw", True)
        Call updateVisibilityOfDropdowns("drsdgw", True)
        Call updateVisibilityOfDropdowns("drmgw", True)
    ElseIf dataSource = "ST" Then
        Call updateVisibilityOfDropdowns("drdst", True)
        Call updateVisibilityOfDropdowns("drsdst", True)
        Call updateVisibilityOfDropdowns("drmst", True)
    ElseIf dataSource = "FA" Then
        Call updateVisibilityOfDropdowns("drdfa", True)
        Call updateVisibilityOfDropdowns("drsdfa", True)
        Call updateVisibilityOfDropdowns("drmfa", True)
    ElseIf dataSource = "MC" Then
        Call updateVisibilityOfDropdowns("drdmc", True)
        Call updateVisibilityOfDropdowns("drsdmc", True)
        Call updateVisibilityOfDropdowns("drmmc", True)
    ElseIf dataSource = "TA" Then
        Call updateVisibilityOfDropdowns("drdta", True)
        Call updateVisibilityOfDropdowns("drsdta", True)
        Call updateVisibilityOfDropdowns("drmta", True)
    End If
    Range("dimensionsCalc" & varsuffix).Calculate
    Range("metricsCalc" & varsuffix).Calculate
End Sub

Sub eraseReg()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To 10
        Select Case i
        Case 1
            varsuffix = ""
            dataSource = "GA"
        Case 2
            varsuffix = "AW"
            dataSource = "AW"
        Case 3
            varsuffix = "FB"
            dataSource = "FB"
        Case 4
            varsuffix = "AC"
            dataSource = "AC"
        Case 5
            varsuffix = "YT"
            dataSource = "YT"
        Case 6
            varsuffix = "FA"
            dataSource = "FA"
        Case 7
            varsuffix = "TW"
            dataSource = "TW"
        Case 8
            varsuffix = "ST"
            dataSource = "ST"
        Case 9
            varsuffix = "MC"
            dataSource = "MC"
        Case 10
            varsuffix = "TA"
            dataSource = "TA"
        End Select
        Call RegKeySave("HKLM\SOFTWARE\Supermetrics\SMDG\DE" & varsuffix, " ")
        DeleteSetting "Supermetrics", "SMDG", "DE" & varsuffix

        Debug.Print "set: " & GetSetting("Supermetrics", "SMDG", "DE" & varsuffix)
        Debug.Print "reg: " & RegKeyRead("HKLM\SOFTWARE\Supermetrics\SMDG\DE" & varsuffix)
        Range("DE" & varsuffix).value = ""
        Range("dexpired" & varsuffix).value = False
    Next i
End Sub


