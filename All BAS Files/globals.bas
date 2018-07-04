Attribute VB_Name = "globals"
Option Explicit
Public Const columnLimit As Long = 16384
Public Const rowLimit As Long = 1048576


'Excel 2007 ->
'Public Const columnLimit As Long = 16384
'Public Const rowLimit As Long = 1048576

'Excel 2003
'Public Const columnLimit As Long = 256
'Public Const rowLimit As Long = 65536

Public useProxy As Boolean
Public useProxyWithCredentials As Boolean
Public proxyAddress As String
Public proxyUsername As String
Public proxyPassword As String
Public excelVersion As Long
Public usingMacOSX As Boolean
Public useMSXMLforParsing As Boolean
Public useQTforDataFetch As Boolean
Public randID As Long
Public clientLoginModeForGA As Boolean
Public appName As String
Public MSXMLversion As Integer

Public creatingClientFiles As Boolean

Public profileSelectionsArr As Variant
Public runningMultipleReports As Boolean

Public OSandExcelVersion As String
Public calculationSetting As Integer




Public demoVersion As Boolean
Public licenseStatus As String
Public demoStatus As String
Public licenseDaysLeft As Variant
Public stParam1 As String
Public stParam2 As String
Public stParam3 As String
Public stParam4 As String
Public prevCL As String

Public sheetProtectionRemoved As Boolean

Public rCount As Long

Public varsuffix As String

Public Const configSheetBackgroundColorIndex As Integer = 2

Public Const buttonColour As Long = 16777215  '15658734 ' 14277081   'RGB(217, 217, 217)
Public Const buttonBorderColour As Long = 15658734    '10921638   'rgb(166,166,166)
Public Const buttonFontColorIndex As Integer = 1
Public Const buttonFontColor As Long = 4210752    'RGB(64, 64, 64)


Public Const buttonColourLight As Long = 15921906   'rgb(242,242,242)
Public Const buttonBorderColourLight As Long = 14277081   'RGB(217, 217, 217)

Public Const buttonColourRed As Long = 16777215    ' 15658734 '53221    '6382079   'RGB(255, 97, 97)

Public Const chartSeriesBlue As Long = 10904832    '13395456   'RGB(0, 102, 204)

Public parameterColumnOffset As Long
Public Const reportStartColumn As Integer = 3

Public segmDimValuesArr() As Variant
Public segmDimCategoryArr() As Variant
Public segmDimIncludesYear As Boolean
Public segmDimIncludesMonth As Boolean
Public segmDimIncludesWeek As Boolean
Public segmDimIncludesDate As Boolean
Public segmDimNumForYear As Integer
Public segmDimNumForWeek As Integer
Public segmDimNumForMonth As Integer
Public segmDimNumForDate As Integer
Public metricsCountInclSub As Long
Public subDimensionCountOrigForLastDim As Integer
Public dimensionsCombined As String
Public firstMetricStr As String
Public weekType As String

Public reportRunType As String
Public sumAllProfiles As Boolean
Public allProfilesInOneQuery As Boolean
Public allProfilesStr As String
Public rawDataReport As Boolean

Public maxParameterValue As Long  'combined max for iterationnum, segmentnum, metricsetnum

Public arrRivi As Long

Public separeteQueryForEachGAMetric As Boolean

Public comparisonType As String
Public linkedPPT As Boolean

Public profID As String
Public profNum As Long
Public profileCount As Long

Public reportRunSuccessful As Boolean
Public refreshBeforeExport As Boolean
Public reportsFound As Boolean
Public PPTtemplatePath As String

Public foundDimValuesArr As Variant

Public mostGranularTimeDimension As String
Public showNoteStr As String
Public doActionStr As String

Public aika As Long
Public aika1 As Long

Public giveMaxResultsWarning As Boolean
Public giveUniqueSumWarning As Boolean

Public advancedSettings As String

Public filterUFArr As Variant
Public filternum As Integer

Public dateDimensionIncluded As Boolean

Public dateRangeType As String
Public dateRangeTypeDisp As String
Public timePeriodForXdateRange As String

Public selectedDate As String
Public selectDatesLabel As String

Public columnInfoArr As Variant
Public folderDimensionIncluded As Boolean
Public dimensionsRequiringCompressionInSD As Boolean
Public columnModificationsStr As String

Public loginInfoCol As Integer

Public dimensionsBasicStr As String

Public nameEncodingStr As String

Public clientCustomerID As String

Public appID As String
Public serviceName As String
Public moduleName As String

Public varsSheetForDataSource As Worksheet

Public questionUFb1Clicked As Boolean
Public questionUFb2Clicked As Boolean
Public questionUFtextFieldValue As String

Public storeWB As Workbook
Public storeWBlastRow As Long
Public storeWBsheet As Worksheet

Public loopTimer As Double
Public loopIterationCount As Long

Public email As String
Public password As String
Public usernameDisp As String
Public developerToken As String
Public emailLastCheckedOK As Date
Public licenseType As String
Public loginType As String

Public configsheet As Worksheet
Public profileListStart As Range
Public referToProfilesAs As String
Public referToProfilesAsSing As String
Public referToAccountsAsSing As String

Public visibleMetricColumnsCount As Long

Public progresspct As Long

Public doTotals As Boolean
Public doHyperlinks As Boolean

Public accountID As String
Public accountArr As Variant

Public maxSimultaneousQueries As Long

Public Const maxSimultaneousQueriesGA As Long = 19
Public Const maxSimultaneousQueriesAW As Long = 19
Public Const maxSimultaneousQueriesAC As Long = 19
Public Const maxSimultaneousQueriesFB As Long = 19
Public Const maxSimultaneousQueriesYT As Long = 19
Public Const maxSimultaneousQueriesFA As Long = 19
Public Const maxSimultaneousQueriesGW As Long = 19
Public Const maxSimultaneousQueriesST As Long = 19
Public Const defaultMaxResultsPerQuery As Long = 10000

Public initialFetchRound As Boolean
Public combinedRequestStr As String

Public reportContainsSampledData As Boolean

Public separatorList As String
Public Const rscL0 As String = "!"
Public Const rscL1 As String = "#"
Public Const rscL2 As String = "`"
'Public Const rscL2 As String = "«"
Public Const rscL3 As String = "^"
Public Const rscL4 As String = "~"

Public fontName As String

Public authToken As String

Public dataSource As String

Public queryNum As Long
Public queryCount As Long
Public objHTTPnum As Long
Public objHTTParr As Variant
Public queryArr As Variant
Public SDlabelsQuery As Boolean
Public SDothersQuery As Boolean
Public querynNumOfPrevSDlabelsQuery As Long
Public columnModificationsArr As Variant
Public dimCountColumnsArr As Variant
Public dimensionCountMetricIncluded As Boolean
Public vriviChart As Long
Public dimensionHeadersCombined As String

Public allQueriesStarted As Boolean
Public allQueriesFetched As Boolean
Public allXMLsParsed As Boolean

Public queryURL As String
Public objhttp As Object


Public HTTPrequestType As String
Public HTTPpostStr As String

Public errorsFoundInArr As String
Public tokenRefreshed As Boolean

Public foundNonFinishedQuery As Boolean
Public queriesCompletedCount As Long
Public queryReady As Boolean
Public pctDonePrev As Long


Public Const keywordCol As Integer = 5
Public Const errorNoteCol As Integer = 6
Public Const tempCol As Integer = 7
Public fetchKeywordIdeas As Boolean
Public resultSheet As Worksheet

Public profIDRow As Integer
Public profNameRow As Integer
Public accountNameRow As Integer
Public metricNameRow As Integer
Public segmDimRow As Integer
Public segmentRow As Integer
Public firstHeaderRow As Integer
Public lastHeaderRow As Integer

Public segmDimCategoriesCount As Long
Public segmDimCategoryNum As Long
Public segmDimCount As Long
Public queryType As String
Public groupByMetric As Boolean
Public runningSheetRefresh As Boolean
Public importingFromOldVersion As Boolean
Public arr As Variant
Public metrics As String

Public metricsArr(1 To 22, 1 To 14) As Variant
'1 name disp
'2 name
'3 metrics list
'4 metrics count
'5 operation
'6 formatting
'7 basemetric
'8 invert condform
'9 dimension for dimcount metrics
'10 goal number for goal name fetch
'11 metricSetNum
'12 dont calculate totals (implemented for AW only)
'13 datasource
'14 profID

Public metricSetNum As Long
Public metricSetsCount As Long
Public metricNumInMetricSet As Long
Public metricNumInclSubInMetricSet As Long
Public metricSetsArr As Variant
Public metricSetsPrecedingSubMetricCountsArr As Variant

Public metricsCount As Long
Public metricNum As Long
Public dimensions As String
Public dimensionsArr(1 To 10, 1 To 4) As Variant
Public dimensionsCount As Long
Public dimensionsCountInclSubGlobal As Long
Public dimensionsCombinedCol As Long
Public extraDimensionColumns As Long
Public postConcatDimensionIncluded As Boolean
Public firstMetricCol As Long
Public timeDimensionIncluded As Boolean
Public nonTimeDimensionIncluded As Boolean
Public deleteEmptyColumns As Boolean    'whether to delete empty columns in segmented dimension queries - if set to FALSE they will be hidden
Public profilesArr As Variant
Public tempSheet As Object
Public tempArr As Variant
Public useFilterList As Boolean
Public sortType As String
Public sortStr As String
Public createCharts As Boolean
Public buttonObj As Object
Public dataSheet As Object
Public vriviData As Long
Public vsarData As Long
Public sortingCol As Long
Public goalsIncluded As Boolean
Public goalsArr As Variant
Public profilesStr As String
Public segmentsStr As String
Public goalsStr As String


Public maxResults As Long
Public maxResults1 As Long
Public maxResultsMultiplierForComparisonQuery As Long
Public givemaxResultsPerQueryWarning As Boolean
Public avoidSampling As Boolean
Public filterStr As String
Public segmentID As String
Public segmentNum As Long
Public segmentCount As Long
Public segmentName As String
Public segmentIsAllVisits As Boolean
Public segmentArr As Variant
Public startDate As Variant
Public endDate As Variant
Public startDate1 As Variant    'actual reporting period
Public endDate1 As Variant    'actual reporting period
Public startDate2 As Variant    'comparisons to earlier period
Public endDate2 As Variant    'comparisons to earlier period
Public comparisonValueType As String
Public iterationsCount As Long
Public iterationNum As Long
Public updatingPreviouslyCreatedSheet As Boolean
Public sheetName As String
Public sheetID As String
Public segmDimNameDisp As String
Public segmDimName As String
Public segmDimName2 As String
Public segmDimNameCombDisp As String
Public segmDimNameComb As String
Public resultStartRow As Long
Public resultStartColumn As Long
Public resultStart As Range
Public Const buttonSpaceBetween As Integer = 2
Public Const buttonTop As Integer = 5
Public Const buttonWidth As Integer = 50
Public Const buttonHeight As Integer = 30
Public segmDimHasNonTimeComponent As Boolean
Public segmDimIsTime As Boolean
Public doComparisons As Long
Public dateRangeRestriction As String
Public includeOther As Boolean


Public processIDsStr As String
Public processStatusTimer As Double
Public processQueriesCompleted As Long
Public processQueriesTotal As Long
Public objHTTPstatus As Object
Public objHTTPstatusRunning As Boolean
Public inDataFetchLoop As Boolean
Public timerArr As Variant
Public timerStart As Double

Public numberOfCharsThatCanBeReturnedToCell As Long
