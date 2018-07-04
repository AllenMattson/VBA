Attribute VB_Name = "controlMacros"
Option Private Module
Option Explicit

'Versions
'1.00
'1.01 Fix for Peter H: Compile error in hidden module
'1.02 Time dim notifier
'1.10 New version for Scott G: segmenting dimension, refreshing queries, progress bar, chart type toggle etc
'1.11 Conditional formatting
'1.12 Max results exceeded warning
'1.13 Authentication failed warnings
'1.14 Fix SD query columns needed calc, store segmdimnamedisp
'1.15 Use sheet IDs instead of names
'1.16 Fix case-insensitivity issue in SD queries
'1.165 Fetch sheet id by name from storage if deleted, refresh all queries changed
'1.166 Demo date constants
'1.167 Email check case-insensitivity
'1.168 More categories for segm dim, fix column check calculation bug
'1.20  Adwords API, new UI, two segmenting dimensions, proxy settings
'1.201 Looping queries to increase max results
'1.202 Adwords fixes for truncate array, UI colour changes
'1.203 sortingCol fix
'1.204 Macro warning sheet
'1.205 Instructions for proxy settings
'1.206 Filter string fix for aggregate query
'1.207 Filter string change (remove spaces if not found), better error handling for getGAdata and getGAaccountData (MSXML6) - this version had a serious bug
'1.208 Increase getGAdata timeout
'1.209 Adwords positionssum with time dimension fix
'1.210 Excel 2010 compatibility issues - serious bug fixed
'1.211 Xpath syntax for msxml6, using local-name()
'1.212 Adwords cost estimator
'1.213 Disable manual calculation
'1.214 AW client fetching for nested MCCs
'1.215 Using AW API v201008
'1.216 Protecting sheets
'1.217 Protecting sheets improvements
'1.218 Scatter plot copy on error resume next
'1.219 Fix catching illegal combination error
'1.220 Unhide settings when creating chart, deleting sheet
'1.221 Select charttemplate sheet before copying chart
'1.222 AW Date dimension
'1.223 Digital signature, macro instructions, input boxes formatting
'1.224 Correcting query path for AW report definition error string
'1.225 Select chart before copy
'1.226 Chart series colour for Excel 2003
'1.227 AdWords demo capability
'1.228 GA Management API, fetch Adwords non-MCC account if used as login account, faster profile list colouring
'1.229 Segment list fetch fix, faster profile list colouring, CAPTCHA user form, dev token message
'1.230 Change querytype from SD to D if only segmenting dimensions selected
'1.231 PPT generator fixes, deleting sheets fixes, take sheetid first from A1 value, then name
'1.232 Asynchronous profile list fetch
'1.233 QueryTable datafetch & built-in XML parsing
'1.30  Mac compatibility, asynchronous GA data fetch, login screen, logout buttons, improved GA profile list fetch,
'      AW cost metrics micros change, change percentage calculation for time dimensions
'1.301 Proxy settings in GA login box
'1.302 Memory leak handling
'1.303 Resume next bug fix
'1.304 Demo version automatic check
'1.305 Minor comparison date note fix
'1.306 Hide checkboxes and dropdown on logout
'1.307 Profile list OSX fix (unprotect configsheet)
'1.308 OSX fixes (find row UDF etc)
'1.309 Duplicate chart instead of copy, A query chart data range fix
'1.310 Progress meter to statusbar, erase arr before reuse
'1.311 ProgressBox2 shape for OSX
'1.312 Lic check, merge cells OSX fix
'1.313 Option private module
'1.314 Match profile number fix, check box modification
'1.315 Hide sheet before delete
'1.316 segmDimCategoryArr storing in queryarr
'1.317 Note for proxy not working on Mac, optimization, removing unnecessary manual cond formatting
'1.318 Instructions link, note for selecting profiles
'1.319 Various OSX fixes (eg. evaluate calculations), date validation, date formatting
'1.320 Profile selection buttons fix
'1.321 Demo version change to Jul-Oct 2010, stop button change, freeze panes on sheet delete fix, new login buttons
'1.322 New dimensions and metrics (http://analytics.blogspot.com/2011/01/127-new-dimensions-and-metrics.html), refreshcount for dimension queries
'1.323 Profileselections sheet customization fix
'1.324 Replace question mark and asterisk in findrow
'1.325 More search metrics, dimensions and metrics reference link, firstChartToModify
'1.326 Automatic proxy detection, connection testing, AW duplicate query fix
'1.33  Comparison fixes (works better with segm dim), report migration tool
'1.331 Checke params
'1.332 condFormDropDown visibility
'1.333 SSL connect
'1.334 Compare date ranges vs year fix
'1.335 OSX fixes
'1.336 FileNameAndPath datatype fix, importingFromOldVersion
'1.337 Scatterplots for SD queries
'1.3371 Delete data connections on logout, separate objHTTP for auth & control
'1.34  GAQR liccheck, maxSimultaneousQueries = 4, reportContainsSampledData
'1.341 Scatterplot fixes for blanks and NEW rows, AW date formatting fix
'1.3411 Profile list fetch error checking added setQueryAsNotFetched
'1.3412 Don't send same error multiple times to CL
'1.342 MTanalytics SSL, fetchAndStoreAWreportDefinition, AW report definition php
'1.343 Clear login credentials
'1.3431 Direct auth connect on captcha error (Windows&GA)
'1.3432 Credentials encoding fix
'1.344 gaqrawrd4, instructions after profile fetch, encoding fixes, enableevents fixes
'1.345 OSX fixes (use http if https not working)
'1.346 AW API v201101
'1.347 Scatterplot fix for Excel 2003, new macro warning
'1.348 Can handle account feed "Document has moved" error
'1.349 New metrics and dimensions (http://groups.google.com/group/google-analytics-api-notify/browse_thread/thread/be44ee65c450f255/c8816e1215ee0151?show_docid=c8816e1215ee0151&pli=1)
'1.350 AW account fetch uses v201101
'1.351 Mask pw (gaqr4), unprotect sheets before profile fetch dd visibility, disable first screen updates for Mac, location of "Categories shown in charts:"
'1.352 OSX progressbox improvements
'1.353 Configsheet dropdown visibility, querytype A invert condform, AW invert condform, hide empty segmdim label if 2 segmdims
'1.3531 Findrow fix for tilde, question mark and asterisk, filter operator max length to 2 chars
'1.3532 Findrow fix (matching row where value contains tilde, question mark and asterisk)
'1.3533 Faster profile selection
'1.3534 Store demo status to vars sheet and registry
'1.354 Fix for queries that split by month or week and another dimension
'1.355 For OSX, fetch whole profile list with one query
'1.360 AW use MT token, gaqrawrd6, querytype A specified in beginning of aggregate query procedure
'1.3601 More detailed stParam1, uprotect sheets on AW remaining units update, AW time dimension fox when changing querytype SD->D
'1.3602 Report migration fixes, findRangeName function
'1.361 AW token expiration fix, buy license UF, button formatting
'1.3611 More detailed stParam1, AW sheet name fix
'1.3612 AW error message fix
'1.3613 More detailed stParam1, keep value on STOP, AW clientcount = 0 fix
'1.3614 GAQR4 gettoken control improvements
'1.3615 More control, protect sheets on auth error, fix refresh all of adwords sheet
'1.3616 Direct toekn fetch, profiles via MT, gaqrst
'1.370 Calendar UF
'1.3701 SD->D querytype transform fix for dimensionsArr(dimensionsCount, 4)
'1.380 OAuth, PPT checkboxes removed from configsheet
'1.3801 Configsheet fix
'1.3802 Try multiple times to fetch email
'1.381 Automatically detect OAuth authorization on Windows, on Mac test for 2 mins
'1.3811 Don't refresh token when using OAuth, try multiple times to get OAuth headers
'1.382 chrdecode for OSX OAuth headers fetch, select all/none profiles
'1.383 Launch OAuth login when pressing button arrow
'1.384 OSX OAuth fixes, autodetect OAuth authorization in OSX, no registry control on OSX
'1.3841 Querytype A chart placing fix
'1.40 Multiaccount login, OAuth improvements, clear only data columns on sheet refresh
'1.401 Cells locked fix
'1.41 Dynamic date ranges
'1.411 Control improvements (incl. querytype A)
'1.412 OAuth via https, refresh all reports does not override report date settings, loop indicator to AW queries
'1.413 Control improvements, select 1st profile on login
'1.414 Missing charts returned, missing charts notifier added, number format for AW metrics via GA API
'1.415 Broken links error
'1.416 Control improvements
'1.417 Use native uriencode procedure
'1.418 Uriencode try with adodb, then native, added daterangetypes for last week
'1.419 Missing chart title fix, more control
'1.4191 Type conversion (Val) to comparison queries
'1.4192 Check connection to MT Analytics
'1.4193 AW profile list fix, public version of GA functions, protectsheet after prof fetch
'1.4194 Fix for not all previous data being cleared on refresh
'1.4195 AW report def error checking fix (do chrdecode)
'1.420 Page load time metric
'1.421 Dynamic dates for AW reports
'1.422 New dynamic date ranges (last year to date, 2 years, 3 years)
'1.423 Fix for some date ranges having extra day
'1.424 AW query fix, AW client list fetch fix
'1.5   BingAds, gaoaqr2, async AW data fetch
'1.51  Modules sheet, AC automatic report type selection, preserve right side formulas, copy formulas on migration,
'      use QT if MSXML fails, maxSimultaneousQueries = 9 for AW and AC
'1.511 AC Avg position field check fix, remove reference to MSXML, OSX progress AM fix
'1.512 GA login fix (maxsimultanousequeries), Modules sheet buttons for instructions, forum, purchasing
'1.513 Connection error handlers, totals and averages to reports
'1.514 Proxy settings fixes
'1.515 Collision control to A queries, use unrequestedfieldsarr for AW queries with filters, migration tool get dates fix
'1.516 Send failed error handling for runObjHTTP, getEmail error handling, update GA date range on open
'1.517 Migrate non-report sheets, don't warn when finding non-report sheets on all report update, stParam4
'1.52  Multilogin for AW & AC, GA profile fetch with func, Auto Filters
'1.521 Migration tool improvements, findrow case sensitivity fix, improved control, keep autofilters on refresh
'1.53  Search keyword tool, on error resume next for xml parsing scripts
'1.531 Error control for search tool
'1.532 Hide AW units in search tool if own dev token
'1.5321 Button macro fix
'1.533 ReadMe sheet
'1.534 Disable ReadMe sheet, add buttonFC control objects to Modules sheet
'1.535 Migration to new file
'1.536 AW token expired error fix, keyword tool getideas button fix
'1.5361 Note of free API units
'1.5362 Remove sheet class modules, restore sheet and wb event macros
'1.5363 AW report type change fix
'1.5364 RandID
'1.5365 Devtok hide
'1.5366 CLmode for JK, set manual calc during report run, notify of missing profiles when refreshing
'1.5367 Control, isarray check
'1.54  Directory levels, erase segmDimCategoryArr fix
'1.541 Strip # from directory level
'1.5411 Erase stparam4 on procedure exit
'1.542 Sort button
'1.5421 Login process isarray test
'1.55 Declarations, split big modules, clean modules, gaoard
'1.551 querynNumOfPrevSDlabelsQuery fix
'1.5511 DEREG fix
'1.552 Profile & segment fetch error checking, ACDG mode
'1.5521 buttonFCAW fix
'1.5522 Added AW client id -1 error message
'1.5523 ACDG welcome message
'1.56 OAuth 2
'1.561 OAuth expiration check
'1.562 OAuth refresh fixes, license check for getEmail and refreshToken, dimensionCount metrics
'1.563 OAuth 2 for Mac fix, condform compile error fix
'1.564 OSX compilation fixes: split SD procedure, displayruler
'1.565 Login fixes
'1.566 Refresh all button fix, checke email per profile & remove general check
'1.567 AW reports with clienEmail when ID = -1
'1.57 OSX progress box iteration boxes and stop button, remove check for email or password blank on report run, OSX dimensionsCountCurrentXML
'1.58 More than 10000 rows of dimensionCount metrics, fox for > 10000 row queries, all calculated metrics stored first into two sheets
'1.581 Avg visits & days to transaction, AW attempt to fix client id -1, better token handling
'1.582 AW client ID -1 fixed (set clientEmail header on report DL)
'1.5821 Minor fixes, login name greyd out if just one
'1.5822 Hide options
'1.5823 No automatic oauth redirect for Mac
'1.5824 Logos
'1.5825 Feedback & Twitter links
'1.5826 Hide controls on hideOptions
'1.5827 Don't run hideOptions on AW & AC
'1.5828 Remove hyperlink from purchase button
'1.5829 Change to macro message, temp remove cert
'1.583 ACLogin sheet, link to filter instructions
'1.5831 Hideprogressbox on connection error, Mac-specific error warnings, fix 'Logged in with account' on Modules
'1.59 AW API v201109, dev token & API units not needed
'1.591 AW automatic report type selection
'1.592 Splitting splitByDimensions to 3 parts, remove unused variables
'1.593 Macro instructions button, better error message for no accounts found
'1.594 AW filters fixes
'1.5941 More AW filter fixes
'1.5942 AW SD query fix
'1.5943 AW captcha URL fix, impression share as metric
'1.60 FB module, new filtering, Excel & PDF export, mass export & delete, GA data API 3.0
'1.601 FB wall posts fix
'1.602 OSX shape position fix, other OSX fixes, fix for complex metrics, cl3.php
'1.6021 Reg check fix
'1.6022 HTTP error fix, FB fix for queries with only wall post dim
'1.6023 Send rid, hide names
'1.6024 Lock FB cells
'1.6025 OSX progressbox fixes
'1.6026 Checke changes, send demo expired
'1.6027 Custom date range fix, profile selection ix for multi-report run, migration tool fixes
'1.603 Declare variants, function parameters & results
'1.604 Excel append on export, special settings from cl3
'1.605 Faster dimcount metrics calc, don't allow combining
'1.6051 Autologin fixes, chart data series name updating, updates number of visible chart data series
'1.6052 Refresh token on 401 error
'1.6053 Preserve formatting fixes
'1.6054 Scatterplot fixes
'1.6055 GA goal metrics all goals
'1.6056 Hide stop button in beginning of OAuth
'1.606 Fetch account data with email on OAuth (gaoa4_getaccount.php, gaoa4_getemail.php)
'1.6061 Set GA Max results to 10000
'1.607 OSX filterUF fixes   |    gaR1.php around this time (Dec 2011)
'1.6071 OSX filterUF unload fix
'1.61 Metric sets
'1.6101 Placing automation buttons on config sheets
'1.6102 Fix for metricsets with modules other than GA
'1.611 Interface changes
'1.6111 2003 fixes (scatterplot, options blocker)
'1.612 Welcome message changes, preload GA sample query
'1.613 Fix for clearing advanced segment
'1.6131 Profile list colour fix, logout one account fix, objhttps to module level variables
'1.6132 Warning message font size & colour
'1.614 Word count in keyword
'1.6141 Goals control
'1.6142 AW non-MCC account login fix (can use ASK_FOR_ID parameter)
'1.6143 Clear filters on logout
'1.6144 Fix for preserve rigth side values on update
'1.6145 Delete cond formats on update
'1.6146 Modules sheet UI colours, buttonFC visible on logout
'1.6147 ASK_FOR_ID parameter fix
'1.615 Select year in calendar, Modules sheet UI improvements
'1.6151 Fix for queries split by hour comparing to previous hour
'1.6152 Returning visits & visitors metrics
'1.6153 Deleted blank rows from metrics list, control
'1.6154 Timeout on shorten URL
'1.616 New site speed metrics
'1.617 Fix for avg days/visits to transaction
'1.618 AW MCC mapping fix, TotalConvValue fix (trust column positions in AW/AC response), auto adjust mac results to 5000000 with dimcount metrics
'1.6181 Fix for goal numbers, goal starts
'1.6182 Unhide button for AW dev token input
'1.6183 Fixes for 2nd series
'1.62 Last x date range types, GA day of month, dates in querystorage, FB relogin prompt on token expiry, use full PPT range, comparisons *, subtotal formulas
'1.6201 Fix for FB SD queries
'1.6202 FB login fix: datasource from var
'1.6203 GA calculate AW ratio metrics from abs values
'1.621 Comparison type (%, abs, value), fix for FB grouping, allow multiple subqueries with GA directory level dimension, set sort type in querystorage
'1.6211 Keyword tool auth fix, remove mso constants and ref to office object library, fix for specifying date range in qs
'1.6212 Fixes to report migration tool
'1.6213 Sorting fix, comp number format fix
'1.6214 Fix for 2nd metric axis title, legend size on refresh (take chartnum from chart name var N), faster profile check on report run
'1.6215 Remove link, sort fix for A and SD->D queries
'1.62151 Longer timeout on account fetch
'1.6216 Ecommerce conversion rate added
'1.622 Login improements for large accounts (incl. light encoding)
'1.623 WeekISO for GA, refresh hidden sheets, light encoding fix
'1.6231 WeekISO year number fix
'1.6232 Last X ISO weeks data range type
'1.6233 Fix for large account fetch (arrayType=combined)
'1.624 Match metricsetnum when marking subqueries as fetched after empty results, goal names header fix
'1.6241 OSX fix for prof names, faster array compression function
'1.62421 Fetch large profile str in pieces
'1.6243 Add date to oauthrequestid
'1.6244 Report migration fixes for dates and sorting
'1.6245 Check sheetID starts with _SH on refresh all reports
'1.6246 No freeze panes on refresh
'1.6247 System paramto queries, resume next on run query from qs
'1.6248 Multiaccount fetch: max 18 accounts per iteration
'1.6249 Fix for charts with one row expanding on refresh
'1.62491 OSX progress box position relative to visible area
'1.6250 Option to log out of all GADG files & destroy tokens (GA & FB)
'1.62501 Log out all only on user-initiated logout
'1.6251 OSX utf8 decode on QT
'1.626 Two-step OSX data fetch, profile list rowLimit fix, SD chart update fix
'1.6261 Re-use QT sheet
'1.62611 Set rscLs dynamically
'1.62612 Update progress fix for fetching profs rscL4
'1.62613 Show email during Oauth
'1.6262 URL encode for OS X data
'1.6263 Progresspct > 25 on profile fetch retry, don't colour change pcts when docolours = false
'1.627 Visit bounce rate, visit length dim, check for MSXML version, don't check internet connection on report run start
'1.6271 queryIDforDB use rscl1, AW totalconversionvalue format to d1000000
'1.62711 Fix GA login button macro name
'1.63 Multi-segment reports, QT urldecode fixes (decode when swith to http, counter datatype)
'1.6301 Sortingcol fix for D/SD
'1.6302 AC v8, deleted deprecated FB metrics
'1.6303 AC parse custom error msg from login response
'1.6304 AC use acAccountF.php
'1.6305 change segmentSelectionCodes number format
'1.631 Get token with getemail, profile fetch arrtype combined2
'1.632 UF images as bmp, empty objhttp on query rerun, error msg placement on querytype A
'1.6321 Separate sortStr for each metricset, separeteQueryForEachGAMetric setting
'1.63211 Segments in alphabetic order
'1.633 Use gaR1ms for OS X GA data fetch
'1.64 New GA fields https://groups.google.com/forum/#!topic/ga-dev-changelogs-reporting-core/5fjkrJOIVOI/discussion, don't allow combining compressed dimensions with unique metrics
'1.6401 report migrate check if config sheet before formatting, set separeteQueryForEachGAMetric on qs
'1.641 Modify query button, remove PPT settings from qs
'1.642 AW API v201206 report download as TSV
'1.643 AC module name to Bing Ads, AC data fetch bug fixes (column headers, two time dims), check OS on AC/AW field check
'1.6431 Check for 500 server error
'1.6432 Check for MSXML versions 1-6, fix for convertrscl
'1.6433 Twitter account change to GADataGrabber
'1.6434 setMSXML: setOption 2, 13056
'1.644 Keyword tool to use v201209, OSX progress box in light grey
'1.6441 Ext trial promo
'1.645 mainphp
'1.6451 Fix gaR1ms -> mainph
'1.6452 KW tool ph fname (remove test), equal rowheight for profiles
'1.65 AW Oauth, AW mainph, AW new IS metrics, metric formatting in ph
'1.6501 Clean
'1.6502 Yearweek for AW, combineDimensionLabels use isTime
'1.6503 AC post-query filter fix
'1.6504 Proxy conf fixes: recheck if MSXML has become available
'1.6505 KW tool refresh Oauth token fix
'1.6506 FB yearmonth & yearweek
'1.651 Do segment sort server-side, disable progbox on OSX segm refresh
'1.6511 Remove unused UFs (AW login, captcha), show username on buy license box
'1.6512 gaR1ms for AW, remove subqueries, change rscL2 from |
'1.66 YouTube Module
'1.6601 Hide YT SD sel on login, YT metrics sort
'1.6602 YT metrics fix
'1.6603 YT refresh metric selection fix
'1.6604 YT time watched num format fix
'1.6605 Sheet protection fixes for Excel 2013
'1.6606 WaitForResponse to prof fetch & oauth
'1.6607 Fixed AW field selection crash bug (repl. -- with _), reorder GA fields
'1.6608 FB wall posts fix
'1.6609 Remove AW & AC units reporting
'1.661 AW total & daily budget, mainph returns cur values instead of micros
'1.6611 AW CPC fix
'1.662 AC mainph, AC budget, remove KW tool, label for sorting col
'1.6621 AC login fixes
'1.6622 Chart series type fix for report refresh
'1.663 Refresh profile list button, maxAccountsPerIteration = 1000
'1.6631 Fix for AW cost per conversion metrics (operation:div)
'1.664 More space for GA profname (wp name added to ph), createdby label change, added public functions
'1.6641 SD labels changes, query count per profile, start queries for each profile first
'1.665 sumAllprofiles, allProfilesInOneQuery
'1.6651 Added latitude adn longitude GA dims
'1.67 Raw data reports
'1.68 Split date range to reduce sampling, QT: chrdecode before URL and UTF8 decode, set maxResultsProfile on request, use yearofisoweek with isoweek
'1.681 Column borders in D/SD reports, filter fix for selecting header _
 '1.69 Flickr Module
'1.6901 Send prof & acc name with requests
'1.6902 Replace cdec with clng
'1.691 SDO, new GA fields (Release 2013-07-09)
'1.6911 FL sum of photo metrics
'1.6912 Fix GA custom dims as SD, nvbpw
'1.6913 getProxySettingsIfNeeded before segment refresh
'1.692 FB mainph, SM in all urls except FB OAuth
'1.7 Supermetrics
'1.7001 Blue top bar
'1.7002 Font fixes: prevent overflow on Arial
'1.7003 Export all reports UF PPT/PDF color fix
'1.701 Pages in loginstatusbox, multi checke, send email & appid with purchase link, no msgbox after refresh all
'1.7011 Refresh pivot tables macro
'1.7012 demoNote links
'1.7013 Purchase links to form, goto for anchor
'1.7014 Determine main font
'1.7015 Oauth via SM WP
'1.7016 Alphabetic desc sort, scatter plot marker colour
'1.702 AdSense metrics in GA
'1.7021 FB login via sm domain
'1.7022 Removed font vars from thisworkbook that caused compile error in 64 bit
'1.7023 Wider profname col, tooltip for prof range
'1.703 Format table button for D/SD
'1.7031 Format table button for A, moved useQTforDataFetch and pptChartsPerSlide to Setting sheet, use fontName in PPT, retry nodata just once
'1.7032 Reg control SM
'1.70321 License expiration message URL fix SM
'1.7033 Faster check profiles
'1.70331 pctDonePrev
'1.70332 findRowWithValue change (isnumeric), placeEachResultOnNewRow for first query on D, charttitle to plotarea innerleft
'1.704 OSX dropdown fix with 3d, update YT dates on launch
'1.7041 calculationSetting
'1.70411 Help links point to SM
'1.7042 GA custom row limits, escape special chars in filter value
'1.705 AC OAuth, FB metricExplanation formatting fix
'1.706 Changes to queryManagement, storeEmailCheckedDateToSheet & storeTokenToSheet (use arrs instead of ranges), GA customvarnames, don't wait on token refresh
'1.7061 Regex filtering for modules other than GA
'1.7062 getFromCurrentQuery calc fix for modules other than GA
'1.70621 AC City dim, getTokenFromSheet get first if no match
'1.7063 Demographic dims added to GA, hid isTablet and isMobile
'1.7064 Changed auto adjust max results to 500 000 with dimcount metrics (from 5 000 000), apply this to dimcount metric queries only, added dimcount included flag to metricsetsarr and queryarr
'1.71 Profile list filtering, store prof & acc name in tokenssh, get profile order from querystorage, getToday func
'1.7101 publicMacros, Functions v 2.3 to publicFunctions
'1.72 Twitter Module
'1.7201 TW module OSX support
'1.7202 TW fixes (use chrencode on OSX)
'1.7203 Refresh segment list note, TW search term cell type to text (accepts @ as first char)
'1.7204 refreshDataOnAllSheetsAndPivotTables macro order fix, updatePivotTables setting
'1.7205 convertRSCL on allprofilesstr
'1.721 MCF metrics
'1.7211 Case-insensitive findRowWithValue, only show No data found in raw data reports if no data found for any segment/profile
'1.722 GA checkfields, OR combinations for filters for ds other than GA
'1.7221 GA don't do checkfields if no mcf fields
'1.7222 Save your with DE reg setting
'1.7223 No field check on OSX
'1.7724 checkFields buttons for GA, AC & AW
'1.7725 Case-sensitive findRow
'1.7726 Oauth show error immediately, protectSheets on checkfields
'1.7727 PPT gen: select sheet & chart before copy
'1.7728 Check whether to send account names & tokens for each prof, encoding, max length for combinedRequestStr and run again for remaining
'1.7729 Faster logoutOneAccount, don't add profileselection CBs if profilecount exceeds 2000
'1.773 AW shopping dims, GA nth dims
'1.7731 Logged in error fix (check configsheet visibility), OS X progress box fix for login status box launch
'1.78 GW Module
'1.7801 Send email qith query
'1.7802 GW check demo license fix
'1.7803 Send segment ID before name
'1.7804 Don't escape \ in filters, allow non-date values for custom date range (parsed with strtotime), AW conversions and audience dimensions, fix checkfields dimensions
'1.7805 TW exit fix for logoutOneAccount
'1.7806 Added AW labels
'1.7807 GA fields: content groups, isTrueViewVideoAd
'1.7808 Allow formulas to be used for all parameters in qs (copy values in refreshDataOnSelectedSheet)
'1.78083 set dexpiredTW on setupForRelease
'1.781 Tweet link in progressbox, showalldata on login, enhanced ecommerce, related products & DCM fields
'1.782 YT ads & annotation metrics, remove YT special field selection stuff
'1.7821 Raw reports don't count sorting col into data range, removed GA metrics entrance bounce rate & return user visits
'1.7822 OAuth open page with http if ssl fails (due to Yosemite problems)
'1.79 Stripe module, don't send system param with Oauth URL due to tinyurl problems
'1.791 AW conversion metric name changes, ST avg. metrics for charges, check for very small abs values on SD column hide
'1.8 FB Ads, AW new ad extensions, remove FL, forum link to supermetrics.com/forum
'1.8001 FA drmfa_1 fix
'1.8002 FA prompt relogin on "access token expired" error
'1.801 FA action type metrics & CPA & ROAS, FB positive & negative feedback metrics
'1.8011 Don't delete FB and FA sheets on report delete (issheetaconfigsheet incl those names)
'1.8012 uriEncode(allProfilesStr)
'1.802 AW location type, active view, unique cookies
'1.8021  dontMoveShapesWithCellsAllConfigSh
'1.803 YT traffic source detail, GA keyword match type
'1.804 goo.gl, FB post dimensions, AW assisted conversions & bounce rate, remove FB metric explanations & show wall posts, FB allow more than 1 dim
'1.8041 FB remove positive action rate, added max categories 1000
'1.8042 FA relevance score metric and dims
'1.8043 Hide GW from Modules sheet
'1.805 Advanced settings input for GA & FA
'1.806 Fix for advanced settings when used together with avoid sampling
'1.8061 openOAuthAuthorizationPage use genRandomString for randStr
'1.81 GW is back, AC shopping dimensions
'1.811 useAPIURLs, AW fields: estimated total conversions, ad channel type, final URL
'1.812 Checke uses tokens, fixes for Mac Excel 2016: use numberFormatLocal for dates, empty string instead of "General", dropdown controlformat change value in linkedcell
'1.813 AW video fields, interactions, conversions. Remove useAPIURLs var
'1.81301 AW cost per interaction
'1.82 MC Module
'1.82001 GA fix segment & goal ranges in vars
'1.8201 New VBA certificate for SM, with timestamping (http://timestamp.globalsign.com/scripts/timstamp.dll)
'1.821 FA lead ads dimensions, AW call & first position fields, GA datehour
'1.822 FB video metrics
'1.823 FA video metrics, action values by type, ad recall, other; FB post actions and reactions
'1.83 TA Module, don't do lcase in parseVarFromName
'1.8301 GA totalValue, AW Gmail metrics & label level dims, FA CPC (Link)
'1.831 New fields including FA unique actions by type, product ID, region, AW tracking template & custom URL paramerets, AC goals, budget dims and tracking template & custom URL paramerets
'1.832 Changed the label name to filterNoteFB for FacebookInsights

Public Const versionNumber As String = "1.832"

Public Const debugMode As Boolean = False
Public Const forceOSXmode As Boolean = False
Public Const storeResultsInSeparateWB As Boolean = False
Public Const sendMode As Boolean = False



Sub setupForReleaseAndSAVE()

    Call setupForRelease(True)

End Sub

Sub setupForReleaseAndDONTSAVE()

    Call setupForRelease(False)

End Sub


Sub setupForRelease(Optional saveFiles As Boolean = True)
    On Error Resume Next

    Dim n As Name

    Dim filepath As String

    If debugMode = True Then
        MsgBox "DEBUG MODE!"
        End
    End If

    If forceOSXmode = True Then
        MsgBox "FORCE OSX MODE!"
        End
    End If

    If storeResultsInSeparateWB = True Then
        MsgBox "STORE RESULTS MODE!"
        End
    End If


    ThisWorkbook.Save


    If sendMode = True Then MsgBox "SEND MODE!"
    dataSource = "GA"
    Call clearFieldSelections
    dataSource = "AW"
    Call clearFieldSelections
    dataSource = "AC"
    Call clearFieldSelections
    dataSource = "FB"
    Call clearFieldSelections
    dataSource = "YT"
    Call clearFieldSelections
    dataSource = "GW"
    Call clearFieldSelections
    dataSource = "ST"
    Call clearFieldSelections
    dataSource = "FA"
    Call clearFieldSelections
    dataSource = "MC"
    Call clearFieldSelections
        dataSource = "TA"
    Call clearFieldSelections
    Call hideSDselectionsYT
    Call hideFilterYT
    Call deleteSheets
    AdWords.Select
    Range("profileselectionsAW").Cells(1, 1).Select
    BingAds.Select
    Range("profileselectionsAC").Cells(1, 1).Select
    Analytics.Select
    Range("profileselections").Cells(1, 1).Select
    Facebook.Select
    Range("profileselectionsFB").Cells(1, 1).Select
    YouTube.Select
    Range("profileSelectionsYT").Cells(1, 1).Select
    Webmaster.Select
    Range("profileSelectionsGW").Cells(1, 1).Select
    Stripe.Select
    Range("profileSelectionsST").Cells(1, 1).Select
    FacebookAds.Select
    Range("profileSelectionsFA").Cells(1, 1).Select
    MailChimp.Select
    Range("profileSelectionsMC").Cells(1, 1).Select
  TwitterAds.Select
    Range("profileSelectionsTA").Cells(1, 1).Select


    Analytics.OptionButtons("formattedReportOB").value = 1
    AdWords.OptionButtons("formattedReportOB").value = 1
    BingAds.OptionButtons("formattedReportOB").value = 1
    '     facebook.OptionButtons("formattedReportOB").value = 1
    YouTube.OptionButtons("formattedReportOB").value = 1
    Webmaster.OptionButtons("formattedReportOB").value = 1
    Stripe.OptionButtons("formattedReportOB").value = 1
    FacebookAds.OptionButtons("formattedReportOB").value = 1
    MailChimp.OptionButtons("formattedReportOB").value = 1
    TwitterAds.OptionButtons("formattedReportOB").value = 1
    checkReportFormattingOptionsVisibility (Analytics.Name)
    checkReportFormattingOptionsVisibility (AdWords.Name)
    checkReportFormattingOptionsVisibility (BingAds.Name)
    checkReportFormattingOptionsVisibility (YouTube.Name)
    checkReportFormattingOptionsVisibility (Webmaster.Name)
    checkReportFormattingOptionsVisibility (Stripe.Name)
    checkReportFormattingOptionsVisibility (FacebookAds.Name)
    checkReportFormattingOptionsVisibility (MailChimp.Name)
      checkReportFormattingOptionsVisibility (TwitterAds.Name)

    AdWords.Shapes("illegalFieldsWarningAW").Visible = False
    BingAds.Shapes("illegalFieldsWarningAC").Visible = False
    Analytics.Shapes("illegalFieldsWarning").Visible = False
    Webmaster.Shapes("illegalFieldsWarningGW").Visible = False
    YouTube.Shapes("illegalFieldsWarningYT").Visible = False
    MailChimp.Shapes("illegalFieldsWarningMC").Visible = False
      TwitterAds.Shapes("illegalFieldsWarningTA").Visible = False

    Analytics.Shapes("fieldsOKnote").Visible = False
    AdWords.Shapes("fieldsOKnote").Visible = False
    BingAds.Shapes("fieldsOKnote").Visible = False
    Webmaster.Shapes("fieldsOKnote").Visible = False
    YouTube.Shapes("fieldsOKnote").Visible = False
    MailChimp.Shapes("fieldsOKnote").Visible = False
     TwitterAds.Shapes("fieldsOKnote").Visible = False


    Analytics.Shapes("checkFieldsButton").Visible = False
    AdWords.Shapes("checkFieldsButton").Visible = False
    BingAds.Shapes("checkFieldsButton").Visible = False
    Webmaster.Shapes("checkFieldsButton").Visible = False
    YouTube.Shapes("checkFieldsButton").Visible = False
    MailChimp.Shapes("checkFieldsButton").Visible = False
    TwitterAds.Shapes("checkFieldsButton").Visible = False

    Analytics.ShowAllData
    AdWords.ShowAllData
    BingAds.ShowAllData
    Webmaster.ShowAllData
    Facebook.ShowAllData
    YouTube.ShowAllData
    Stripe.ShowAllData
    FacebookAds.ShowAllData
    MailChimp.ShowAllData
    TwitterAds.ShowAllData


    Call logoutGA
    Call logoutAW
    Call logoutAC
    Call logoutFB
    Call logoutYT
    Call logoutTW
    Call logoutGW
    Call logoutST
    Call logoutFA
    Call logoutMC
    Call logoutTA
    Call clearFiltersGA
    Call clearFiltersAW
    Call clearFiltersAC
    Call clearFiltersFB
    Call clearFiltersYT
    Call clearFiltersGW
    Call clearFiltersST
    Call clearFiltersFA
    Call clearFiltersMC
    Call clearFiltersTA
    Call unprotectSheets
    Call clearQueryStorage
    Call showMacroInstructions
    Call clearProfileSelections
    '  Call hideOptions
    Call hideNames
    Call breakLinks
    Call deleteQueryTables
    Call dontMoveShapesWithCellsAllConfigSh

    '  Call hideGWfromModulesSheet

    'reset exportUF settings
    Range("exportUFfiletype").value = "excel"
    Range("exportUFcombined").value = False
    Range("exportUFsave").value = False
    Range("exportUFsavePath").value = vbNullString
    Range("exportUFclose").value = False
    Range("exportUFrefresh").value = False
    Range("exportUFPPTchartsOnly").value = False
    Range("exportUFexcelAppendPath").value = vbNullString
    Range("enableExcelExportAppend").value = True

    'Analytics.DropDowns("comparisonValueTypeDD").Visible = False
    'Range("comparisonValueTypeSelectionNum").value = 1

    Range("sortSelectionNum").value = 1

    Range("useQTforDataFetch").value = False
    Call unprotectSheets
    Call hideAutomationButtons
    Modules.Visible = xlSheetVisible
    Modules.Select
    With Modules
        .Shapes("apptitleBox").TextFrame.Characters.Text = "version " & versionNumber
        .Shapes("buttonFC").Visible = True
        .Shapes("buttonFCAW").Visible = True
        .Shapes("buttonFCAC").Visible = True
        .Shapes("buttonFCFB").Visible = True
        .Shapes("buttonFCYT").Visible = True
        .Shapes("buttonFCGW").Visible = True
        .Shapes("buttonFCST").Visible = True
        .Shapes("buttonFCFA").Visible = True
        .Shapes("buttonFCMC").Visible = True
        .Shapes("buttonFCTA").Visible = True
    End With

    Range("randID").value = vbNullString
    randID = 0
    Range("rcount").value = 0
    Range("clientLoginModeForGA").value = False

    Sheets("Flickr").Visible = xlSheetVeryHidden


    For Each n In ActiveWorkbook.Names
        If Left(n.Name, 3) = "_SH" Then
            n.Delete
        End If
    Next n

    Call protectSheets
    Range("versionNum").value = CSng(versionNumber)
    Range("dexpired").value = False
    Range("dexpiredAW").value = False
    Range("dexpiredAC").value = False
    Range("dexpiredFB").value = False
    Range("dexpiredYT").value = False
    Range("dexpiredFA").value = False
    Range("dexpiredGW").value = False
    Range("dexpiredTW").value = False
    Range("dexpiredST").value = False
    Range("dexpiredMC").value = False
    Range("dexpiredTA").value = False
    Sheets("vars").Visible = xlSheetVeryHidden
    Sheets("varsaw").Visible = xlSheetVeryHidden
    Sheets("varsac").Visible = xlSheetVeryHidden
    Sheets("cred").Visible = xlSheetVeryHidden
    Sheets("tokens").Visible = xlSheetVeryHidden
    Sheets("logins").Visible = xlSheetVeryHidden
    Sheets("codes").Visible = xlSheetVeryHidden
    Sheets("QT").Visible = xlSheetVeryHidden

    Range("linkPPTlast").value = False
    Range("savedQueryCol").EntireColumn.ClearContents

    Sheets("querystorage").Visible = xlSheetHidden
    Sheets("settings").Visible = xlSheetHidden
    Sheets("proxysettings").Visible = xlSheetHidden

    If Sheets("settings").ChartObjects.Count < 3 Then MsgBox "MISSING CHARTS!"
    If Range("proxyaddress").value <> vbNullString Then MsgBox "PROXY SPECIFIED " & Range("proxyaddress").value
    Modules.Select
    Application.EnableEvents = True
    Cells(1, 1).Select
    Call protectSheets

    Application.Calculation = xlAutomatic

    If saveFiles = True Then

        ThisWorkbook.Save

        filepath = ThisWorkbook.Path & "\"
        'ThisWorkbook.CheckCompatibility = False
        Application.DisplayAlerts = False


        With ThisWorkbook.VBProject.VBComponents("globals").CodeModule
            .ReplaceLine 2, "Public Const columnLimit As Long = 256"
            .ReplaceLine 3, "Public Const rowLimit As Long = 65536"
        End With
        '   ThisWorkbook.CheckCompatibility = False
        Application.DisplayAlerts = False
        'ThisWorkbook.SaveAs "GADataGrabber" & Round(100 * Rnd, 0) & ".xls", 56
        Application.DisplayAlerts = True
        ThisWorkbook.SaveAs filepath & "SupermetricsDataGrabber.xls", 56
        '  ThisWorkbook.CheckCompatibility = False

        With ThisWorkbook.VBProject.VBComponents("globals").CodeModule
            .ReplaceLine 2, "Public Const columnLimit As Long = 16384"
            .ReplaceLine 3, "Public Const rowLimit As Long = 1048576"
        End With


        ThisWorkbook.SaveAs filepath & "SupermetricsDataGrabber " & versionNumber & " " & Round(100 * Rnd, 0) & ".xlsm", 52
        ThisWorkbook.SaveCopyAs filepath & "SupermetricsDataGrabber.xlsm"


    End If
    Application.DisplayAlerts = True
End Sub



Sub checkOperatingSystem()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim OSstring As String

    OSstring = Application.OperatingSystem

    If InStr(1, OSstring, "Win") > 0 Then
        usingMacOSX = False
    ElseIf InStr(1, OSstring, "OS X") > 0 Then
        usingMacOSX = True
    ElseIf InStr(1, OSstring, "OSX") > 0 Then
        usingMacOSX = True
    ElseIf InStr(1, OSstring, "Mac") > 0 Then
        usingMacOSX = True
    ElseIf InStr(1, OSstring, "Win") > 0 Then
        usingMacOSX = False
    Else
        usingMacOSX = False
    End If


    If Range("useOSXsettings") = True Then usingMacOSX = True


    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If forceOSXmode = True Then usingMacOSX = True
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


    appName = "Supermetrics Data Grabber"


    Range("usingMacOSX").value = usingMacOSX

    If usingMacOSX = False Then
        If Range("useQTforDataFetch").value = True Then
            useMSXMLforParsing = False
            useQTforDataFetch = True
        Else
            useMSXMLforParsing = True
            useQTforDataFetch = False
        End If
    Else
        useQTforDataFetch = True
        useMSXMLforParsing = False
    End If

    excelVersion = CInt(Left$(Application.Version, 2))

    OSandExcelVersion = OSstring & "|" & excelVersion & "|"
    If useQTforDataFetch = False Then
        OSandExcelVersion = OSstring & "|" & excelVersion & "|MSXML"
    Else
        OSandExcelVersion = OSstring & "|" & excelVersion & "|QT"
    End If

    If randID = 0 Then
        If Range("randID").value = vbNullString Then
            Randomize
            randID = Int((10000 - 0 + 1) * Rnd)
            Range("randID").value = randID
        Else
            randID = Range("randID").value
        End If
    End If
    OSandExcelVersion = OSandExcelVersion & "|" & randID

    OSandExcelVersion = OSandExcelVersion & "|" & fontName

    stParam3 = OSandExcelVersion


End Sub

Private Sub hideVarsSheet()
    Sheets("vars").Visible = xlSheetVeryHidden
    Sheets("varsAW").Visible = xlSheetVeryHidden
End Sub


Public Function licMsg(postStr As String, skipProblems As Boolean) As String
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim URL As String
    Dim licenseResponseStr As String
    Dim objHTTPcontrol As Object
    URL = "https://supermetrics.com/api/checkLicense?responseFormat=RSCL"
    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, postStr)
        licenseResponseStr = queryTableResultStr
        If debugMode = True Then Debug.Print licenseResponseStr
    Else
        Call setMSXML(objHTTPcontrol)
        If useProxy = True Then objHTTPcontrol.setProxy 2, proxyAddress
        If skipProblems = True Then
            objHTTPcontrol.Open "POST", URL, True
        Else
            objHTTPcontrol.Open "POST", URL, False
        End If
        If useProxyWithCredentials = True Then objHTTPcontrol.setProxyCredentials proxyUsername, proxyPassword
        objHTTPcontrol.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPcontrol.setTimeouts 1000000, 1000000, 1000000, 1000000
        objHTTPcontrol.setOption 2, 13056
        objHTTPcontrol.send (postStr)
        If skipProblems = True Then
            Set objHTTPcontrol = Nothing
            Exit Function
        End If
        licenseResponseStr = objHTTPcontrol.responsetext
        If debugMode = True Then Debug.Print licenseResponseStr
    End If

    licMsg = licenseResponseStr

End Function


Sub checkE(ByVal email As String, Optional ds As String = "GA", Optional skipProblems As Boolean = False, Optional skipMessages As Boolean = False, Optional dontDoLogout As Boolean = False)


    On Error GoTo loppu
    If debugMode = True Then On Error GoTo 0
    dataSource = ds
    Call setDatasourceVariables
    Call checkOperatingSystem
    Call unprotectSheets

    Dim licenseResponseStr As String
    Dim licenseResponseStrComb As String

    Dim postStr As String
    Dim DEreg As String
    Dim DEset As String

    Dim specialStr As String

    Dim serviceName As String

    Dim messageToUser As String
    Dim doAction As String
    Dim tempStr As String
    Dim emailNum As Long
    Dim emailsArr As Variant
    Dim tokensArr As Variant


    If email = vbNullString Then Exit Sub


    emailsArr = Split(email, rscL1)
    ReDim tokensArr(0 To UBound(emailsArr))

    For emailNum = LBound(emailsArr) To UBound(emailsArr)
        tokensArr(emailNum) = getTokenForEmail(emailsArr(emailNum))
        If emailNum = 0 And tokensArr(emailNum) = vbNullString Then tokensArr(emailNum) = authToken   'Let's use the global if no token found. This is the case when this function is called from OAuthLogin before storeTokenToSheet
    Next emailNum

    postStr = "token=" & uriEncode(Join(tokensArr, rscL1)) & "&appid=" & appID & "&version=" & uriEncode(versionNumber) & "&rid=" & randID & "&rscL1=" & uriEncode(rscL1)
    If debugMode Then Debug.Print "POST   " & postStr
    If skipProblems And postStr = prevCL Then Exit Sub

    If skipProblems = False Then licenseStatus = "NOTFOUND"

    licenseDaysLeft = 0

    licenseResponseStrComb = licMsg(postStr, skipProblems)
    prevCL = postStr

    If skipProblems Then Exit Sub



    ReDim tempArr(0 To UBound(emailsArr), 1 To 4)

    For emailNum = LBound(emailsArr) To UBound(emailsArr)
        email = emailsArr(emailNum)
        licenseResponseStr = Split(licenseResponseStrComb, rscL1)(emailNum)

        If email = vbNullString Then
            demoVersion = True
            licenseStatus = "INVALID"
            demoStatus = "INVALID"
        Else
            licenseStatus = parseVarFromStr(licenseResponseStr, "STATUS", "%")
            demoStatus = parseVarFromStr(licenseResponseStr, "DEMOSTATUS", "%")
            licenseDaysLeft = parseVarFromStr(licenseResponseStr, "DAYSLEFT", "%")
        End If

        If licenseStatus = "VALID" Then
            demoVersion = False
        Else
            demoVersion = True
        End If

        tempArr(emailNum, 1) = licenseStatus
        tempArr(emailNum, 2) = demoStatus
        tempArr(emailNum, 3) = licenseDaysLeft
        tempArr(emailNum, 4) = email

        If Not skipMessages Then

            If licenseStatus = "NOTFOUND" Then

                demoVersion = True

                Range("licenseWarningShown" & varsuffix).value = False

                On Error Resume Next
                If usingMacOSX = False Then
                    DEset = GetSetting("Supermetrics", "SMDG", "DE" & varsuffix & CStr(Year(Now)) & "1")
                    DEreg = RegKeyRead("HKLM\SOFTWARE\Supermetrics\SMDG\DE" & varsuffix & CStr(Year(Now)) & "1")
                Else
                    DEreg = vbNullString
                    DEset = GetSetting("Supermetrics", "SMDG", "DE" & varsuffix & CStr(Year(Now)) & "1")
                End If


                If (demoStatus <> "VALID" And demoStatus <> "NOTFOUND") Or Range("dexpired" & varsuffix).value = True Or DEreg = "T" Or DEset = "T" Then
                    '          If demoStatus <> "VALID" Then

                    stParam2 = ""
                    postStr = "email=" & email & "&appid=" & appID & "&version=" & versionNumber & "&p1=DEMO EXPIRED" & "&rid=" & randID
                    If Range("dexpired" & varsuffix).value = True Then stParam2 = stParam2 & "|Rng|"
                    If DEreg = "T" Then stParam2 = stParam2 & "|Reg|"
                    If DEset = "T" Then stParam2 = stParam2 & "|Set|"
                    If demoStatus <> "VALID" Then stParam2 = stParam2 & "|CL|"
                    postStr = postStr & "&p2=" & stParam2
                    tempStr = licMsg(postStr, False)

                    Range("dexpired" & varsuffix).value = True
                    On Error Resume Next
                    If Not usingMacOSX Then
                        Call RegKeySave("HKLM\SOFTWARE\Supermetrics\SMDG\DE" & varsuffix & CStr(Year(Now)), "T")
                    End If
                    SaveSetting "Supermetrics", "SMDG", "DE" & varsuffix & CStr(Year(Now)), "T"
                    If dontDoLogout = False Then Call logoutOneAccount(email)

                    If usernameDisp = vbNullString Then usernameDisp = email
                    Call hideProgressBox
                    With buyLicenseBox
                        .note1.Caption = "Your Supermetrics Data Grabber " & moduleName & " trial for account " & usernameDisp & " has expired. Please visit Supermetrics.com to purchase a license." & vbCrLf & vbCrLf & "When making the purchase on the site, use this user ID:"
                        .usernameTB.Text = email
                        .Show
                    End With
                    End
                End If

                If debugMode = True Then On Error GoTo 0

                With configsheet.Shapes("licenseNote" & varsuffix)
                    .Visible = True
                    .TextFrame.Characters.Text = "Trial days left: " & licenseDaysLeft
                    Modules.Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = "Trial days left: " & licenseDaysLeft
                End With
                If Range("demoversionformatting" & varsuffix).value = False Then Call setDemoVersionFormatting

            ElseIf licenseStatus = "INVALID" Then
                demoVersion = True
                licenseDaysLeft = 0
                If Range("licenseWarningShown" & varsuffix).value = False Then
                    If dontDoLogout = False Then Call logoutOneAccount(email)
                    Call hideProgressBox
                    With buyLicenseBox
                        .note1.Caption = "Your Supermetrics Data Grabber " & moduleName & " license for account " & usernameDisp & " has expired. Please visit Supermetrics.com to purchase a license." & vbCrLf & vbCrLf & "When making the purchase on the site, use this user ID:"
                        .usernameTB.Text = email
                        .Show
                    End With

                    End
                End If
            ElseIf licenseStatus = "VALID" Then
                demoVersion = False
                Range("licenseWarningShown" & varsuffix).value = False
                messageToUser = parseVarFromStr(licenseResponseStr, "MESSAGETOUSER", "%")
                doAction = parseVarFromStr(licenseResponseStr, "ACTION", "%")

                With configsheet.Shapes("licenseNote" & varsuffix)
                    .Visible = True
                    .TextFrame.Characters.Text = "License days left: " & licenseDaysLeft
                    Modules.Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = "License days left: " & licenseDaysLeft
                End With
                If Range("demoversionformatting" & varsuffix).value = True Then Call removeDemoVersionFormatting
            End If

            specialStr = parseVarFromStr(licenseResponseStr, "SPECIAL", "%")
            If specialStr <> vbNullString Then
                If InStr(1, specialStr, "xAEx") > 0 Then Range("enableExcelExportAppend").value = True
                '    If InStr(1, specialStr, "xAbsCompx") > 0 Then Analytics.DropDowns("comparisonValueTypeDD").Visible = True
            End If

            If debugMode = True Then
                Debug.Print "License status: " & licenseStatus
                Debug.Print "License days left: " & licenseDaysLeft
            End If

            If dataSource = "GA" Then Call storeEmailCheckedDateToSheet(email)


            If messageToUser <> "" Then
                MsgBox messageToUser
            End If

            If doAction = "END" Then End

        End If

    Next emailNum

    Exit Sub

loppu:

    Call protectSheets

End Sub
Sub hideNames()
    Dim n As Name
    For Each n In ThisWorkbook.Names
        n.Visible = False
    Next
End Sub
Sub showNames()
    Dim n As Name
    For Each n In ThisWorkbook.Names
        n.Visible = True
    Next
End Sub



Sub setDemoVersionFormatting()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Call unprotectSheets
    Call setDatasourceVariables
    If debugMode = True Then Debug.Print "Setting demo version formatting for datasource " & dataSource
    With configsheet
        .Shapes("apptitle" & varsuffix).TextFrame.Characters.Text = "version " & versionNumber    '& " | " & moduleName    '& " TRIAL"
        .Shapes("demoNote" & varsuffix).Visible = True
        .Shapes("licenseNote" & varsuffix).Visible = True
        Range("demoversionformatting" & varsuffix).value = True
    End With
End Sub

Sub removeDemoVersionFormatting()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Call unprotectSheets
    Call setDatasourceVariables
    If debugMode = True Then Debug.Print "Setting full version formatting for datasource " & dataSource
    With configsheet
        .Shapes("apptitle" & varsuffix).TextFrame.Characters.Text = "version " & versionNumber    '& " | " & moduleName
        .Shapes("demoNote" & varsuffix).Visible = False
        Range("demoversionformatting" & varsuffix).value = False
    End With
End Sub


'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
Sub RegKeySave(i_RegKey As String, _
               i_Value As String, _
               Optional i_Type As String = "REG_SZ")
    Dim myWS As Object
    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'write registry key
    myWS.RegWrite i_RegKey, i_Value, i_Type
End Sub

'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Function RegKeyRead(i_RegKey As String) As String
    Dim myWS As Object

    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    RegKeyRead = myWS.RegRead(i_RegKey)
End Function



