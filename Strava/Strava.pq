/*
Overview

I've tried to limit transformations of data from Strava - this can be done in PowerBI desktop.
This means new fields added by Strava will automatically be available in the data connector.


See  https://developers.strava.com/docs/reference for further information

The Navigation Table in Power BI will show the following Tables:
    Activities (Summary)
    *Activities (Comments)
    *Activities (Detailed)
    *Activities (Detailed) - (Laps)
    *Activities (Detailed) - (Segment Efforts)
    *Activities (Detailed) - (Best Efforts)
    Athlete
    Clubs

 Parameters
    *Start Date &  End Date    

Paging
Paging Parameters:   page & per_page:
    Activities (Summary)    - (Strava ID: List Athlete Activities)
    Clubs                   - (Strava ID: List Club Activities)
            
Paging Parameters: page_size & after_cursor:
    Activities (Comments)   - (Strava ID:  List Activity Comments)
            
No Paging 
    Athlete
    Activities (Detailed)

    


*/

section Strava;
[DataSource.Kind="Strava", Publish="Strava.UI"]
shared Strava.Contents = Value.ReplaceType(StravaImpl, StravaType);


StravaType = type function (
    StartDate as (type datetime meta [
        Documentation.FieldCaption = "Start Date",
        Documentation.FieldDescription = "Actvity Start Date Range"
    ]),
    
    optional EndDate as (type datetime meta [
        Documentation.FieldCaption = "End Date",
        Documentation.FieldDescription = "Activity End Date Range (Optional)"])

    )
    as table meta [
        Documentation.Name = "Strava PBi DataConnector",
        Documentation.LongDescription = "The Power BI Data Connector - wwww.blueoceanweb.co.uk"
    ];




StravaImpl = (StartDate as datetime ,optional EndDate as nullable datetime ) as table =>
    let
        // Start and End date cannot be null so default them
         StartDate2 = if StartDate = null then #datetime( 1980,1,1,0,0,0) else StartDate,
         EndDate2 = if EndDate = null then #datetime( 3000,1,1,0,0,0) else EndDate,

        table = StravaData.NavTable(StartDate2,EndDate2)

        
    in
        table;

shared Strava.ContentsTest = StravaImpl2;

StravaImpl2 = () as table =>
    let

        // Start and End date cannot be null so default them
         StartDate2 =  #datetime( 1980,1,1,0,0,0) ,
         EndDate2 =  #datetime( 3000,1,1,0,0,0) ,

        table = StravaData.NavTable(StartDate2,EndDate2)
        

        
    in
        table;



//Data Source Kind description
Strava = [

    // The first value is the fuction which will be ran when tested.  Shouldn't really use a different function but couldn't get that to work with paranmeters.
   TestConnection = (dataSourcePath) => { "Strava.ContentsTest" },  

    Authentication = [
        OAuth = [                                                               
            StartLogin = StartLogin,
            FinishLogin = FinishLogin                                                     
            ,Refresh = Refresh                                        

            ,Label = Extension.LoadString("AuthenticationLabel")
        ]
    ]
    ,Label = Extension.LoadString("DataSourceLabel")
];


//UI Export definition
Strava.UI = [
    Beta = true,
    ButtonText = { Extension.LoadString("FormulaTitle"), Extension.LoadString("FormulaHelp") },
    SourceImage = Strava.Icons,
    SourceTypeImage = Strava.Icons
];


Strava.Icons = [
    Icon16 = { Extension.Contents("Strava16.png"), Extension.Contents("Strava20.png"), Extension.Contents("Strava24.png"), Extension.Contents("Strava32.png") },
    Icon32 = { Extension.Contents("Strava32.png"), Extension.Contents("Strava40.png"), Extension.Contents("Strava48.png"), Extension.Contents("Strava64.png") }
];



//
// Make sure you have two files called client_id and client_secret in the root of your solution.
// signin to Strava and navigate to https://www.strava.com/settings/api 
// Follow the steps and obtain your client_id and client_secret.
//
client_id = Text.FromBinary(Extension.Contents("client_id"));
client_secret = Text.FromBinary(Extension.Contents("client_secret"));

redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";


// strava base URL
api_uri = "https://www.strava.com/api/v3/";
activity_uri = "athlete/activities";                           
athlete_uri = "athlete";
clubs_uri = "/athlete/clubs";

// default page size, using the MAX page size to limit the amount of calls to the API
per_page = 200;  // numbers of rows per page.  

windowWidth = 1200;
windowHeight = 1000;


/*
https://learn.microsoft.com/en-us/power-query/handling-authentication#implementing-an-oauth-flow

Original OAuth signatures (names are not important sometimes dataSourcePath is resourceUrl)
    StartLogin = (dataSourcePath, state, display) => ...;
    FinishLogin = (context, callbackUri, state) => ...;
    Refresh = (dataSourcePath, refreshToken) =>  ...;
    Logout = (accessToken) => ...;

Advanced OAuth signatures
    All signatures accept a clientApplication record value, which is reserved for future use.
    All signatures accept a dataSourcePath (also referred to as resourceUrl in most samples).
    The Refresh function accepts an oldCredential parameter, which is the previous record returned by your FinishLogin function (or previous call to Refresh).

    StartLogin = (clientApplication, dataSourcePath, state, display) => ...;
    FinishLogin = (clientApplication, dataSourcePath, context, callbackUri, state) => ...;
    Refresh = (clientApplication, dataSourcePath, oldCredential) =>  ...;
    Logout = (clientApplication, dataSourcePath, accessToken) => ...;

*/








StartLogin = (clientApplication, dataSourcePath, state, display) => 
    // Advanced OAuth Signature

        let

         AuthorizeUrl = "https://www.strava.com/oauth/authorize?" & Uri.BuildQueryString([
                client_id = client_id,
                redirect_uri = redirect_uri,
                response_type = "code",
                scope = "read_all,activity:read_all",
                state = state
                ]
                
                )
        in
            [
                LoginUri = AuthorizeUrl,
                CallbackUri = redirect_uri,
                WindowHeight = windowHeight,
                WindowWidth = windowWidth,
                Context = null
            ];



// Advanced OAuth Signature
FinishLogin = (clientApplication, dataSourcePath, context, callbackUri, state) =>       

    let
        Parts = Uri.Parts(callbackUri)[Query]                                                                    
    
    in
        TokenMethod(dataSourcePath,"authorization_code","authorization_code",Parts[code]);      
                                                                                                



Refresh = (dataSourcePath, refreshToken) => 
       // Original OAuth signatures    

    let
        Token = TokenMethod(dataSourcePath, "refresh_token", "refresh_token", refreshToken) 
    in 
        Token;



TokenMethod = (dataSourcePath, grantType, key, value) =>
    // This is called twice, see refresh() and FinishLogin() 
      
    let

        ContentTemp = if grantType = "refresh_token" 
                then 
                    Text.ToBinary(Uri.BuildQueryString([
                                    client_id = client_id,
                                    client_secret = client_secret,
                                    grant_type = grantType,
                                    refresh_token = value                          
                                    ]))
                else  // "authorization_code"
                    Text.ToBinary(Uri.BuildQueryString([
                                    client_id = client_id,
                                    client_secret = client_secret,
                                    grant_type = grantType,
                                    code = value                          
                                    ])),    

        Response = Web.Contents("https://www.strava.com/oauth/token", [                                         
            Content = ContentTemp, 
            Headers=[#"Content-type" = "application/x-www-form-urlencoded",#"Accept" = "application/json"]]), 
        Parts = Json.Document(Response)
    in
        Parts;    


StravaData.NavTable = (optional  StartDate as nullable datetime ,optional EndDate as nullable datetime) as table =>
    let


        // converts text date to a number (Epoch date)
        EndDateAsNumber = Duration.TotalSeconds( EndDate - #datetime(1970,1,1,0,0,0)),         
        StartDateAsNumber = Duration.TotalSeconds(  StartDate - #datetime(1970,1,1,0,0,0)), 
    
        TableClub = StravaData.GetObject(clubs_uri,"Clubs"), 
        TableClubWithActivities = Table.AddColumn(TableClub, "Activities", each getClubActivitires([id])),  
        TableActivities = StravaData.GetObject(activity_uri,"Activities" ),
        
        TableActivities2 = Table.TransformColumnTypes(TableActivities,{{"start_date_local", type datetime}}),
        TableActivitiesWithComments = Table.SelectRows(TableActivities2, 
                each    [comment_count] >= 1 
                    and [start_date_local] >= StartDate 
                    and [start_date_local] <= EndDate),
        TableActivityComments =  getComments(TableActivitiesWithComments) ,
        

        TableActivitiesWithDetActivities = Table.SelectRows(TableActivities2,
                each  [start_date_local] >= StartDate 
                  and [start_date_local] <= EndDate),
        TableDetailedActivity =    getDetailedActivity(TableActivitiesWithDetActivities) ,
            
        
        
        TableDetailedActivityLaps = getTableDetailedActivityLaps(Table.SelectColumns(TableDetailedActivity,{"laps"})),
        TableDetailedActivitySegEff = getTableDetailedActivitySegEff(Table.SelectColumns(TableDetailedActivity,{"segment_efforts"})),
        TableDetailedActivityBestEff = getTableDetailedActivityBestEff(Table.SelectColumns(TableDetailedActivity,{"best_efforts"})),
        TableDetailedActivity2 = Table.RemoveColumns(TableDetailedActivity,{"laps","segment_efforts","best_efforts"}),
        



 source =               #table(            { "Name"         , "Data"}, 
                        {
                        { "Activities (Summary)"        , TableActivities },       
                        { "Activities (Comments)"       , TableActivityComments},
                        { "Activities (Detailed)"       , TableDetailedActivity},
                        { "Activities (Detailed) (Laps)"      , TableDetailedActivityLaps},
                        { "Activities (Detailed) (Segment Efforts)"      , TableDetailedActivitySegEff},
                        { "Activities (Detailed) (Best Efforts)"      , TableDetailedActivityBestEff},
                        { "Clubs"                        , TableClubWithActivities},
                        { "Athlete"                     , StravaData.GetObject(athlete_uri,"Athlete")}                      
                        }
                        
                        ),       

   

        // add other columns
        withItemKind = Table.AddColumn(source, "ItemKind", each "Table"),
        withItemName = Table.AddColumn(withItemKind, "ItemName", each "Table"),
        withIsLeaf = Table.AddColumn(withItemName, "IsLeaf", each true),

        functionTable = AddNavTableFunctions({[Name = "GetObjects", Data = StravaData.GetObject]}),
        
        withFunction = Table.InsertRows(withIsLeaf, Table.RowCount(withIsLeaf) - 1, {[ 
            Name = "Functions",
            Data = functionTable,
            ItemKind = "Folder",
            ItemName = "Folder",
            IsLeaf = false
       ]}),
        
        navTable = Table.ToNavigationTable(withFunction, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf"),
        navTable2 = Table.SelectRows(navTable,each [Name]<>"Functions")

   in 

navTable2;



     


getTableDetailedActivityLaps = (TableDetailedActivity as table) as table =>
let 

    #"Expanded laps" = Table.ExpandListColumn(TableDetailedActivity, "laps"),
    TableDetailedActivityLaps = Table.FromRecords(#"Expanded laps"[laps],null,MissingField.UseNull)

 in
    TableDetailedActivityLaps;

getTableDetailedActivitySegEff = (TableDetailedActivity as table) as table =>
let 

    #"Expanded" = Table.ExpandListColumn(TableDetailedActivity, "segment_efforts"),
    TableRemoveNulls = Table.SelectRows(#"Expanded",each [segment_efforts] <> null),
    TableDetailedActivitySegEff = Table.FromRecords(TableRemoveNulls[segment_efforts],null,MissingField.UseNull)   

 in
    TableDetailedActivitySegEff;



getTableDetailedActivityBestEff = (TableDetailedActivity as table) as table =>
let 

    #"Expanded" = Table.ExpandListColumn(TableDetailedActivity, "best_efforts"),
    TableRemoveNulls = Table.SelectRows(#"Expanded",each [best_efforts] <> null),
    TableDetailedActivityBestEff = Table.FromRecords(TableRemoveNulls[best_efforts],null,MissingField.UseNull)

 in
    TableDetailedActivityBestEff;          

//All of the paging is based off the Tripin Sample https://github.com/Microsoft/DataConnectors/tree/master/samples/TripPin/5-Paging 

// Add a helper function that can be used to retrieve strava API objects 
// that we haven't exposed in the navigation table. We always take a relative path,
// and append it to the base Strava API URL.
// Note: Taking in a full path is a security issue, as it could potentially leak 
// the bearer token to the url specified in the function parameter. 
// For Strava I added an additional parameter to determine the type of content so we can parse this differently later
StravaData.GetObject = (relativeUrl as text, contentType as text, optional StartDate as number , optional EndDate as number)
                            => StravaData.PagedReader(api_uri & relativeUrl, contentType, StartDate, EndDate);		

// Read all pages of data.
// After every page, we check the "NextPage" record on the metadata of the previous request.
// Table.GenerateByPage will keep asking for more pages until we return null.
StravaData.PagedReader = (url as text, contentType as text, optional StartDate as number , optional EndDate as number) =>
    let getNextPage =  ((previous as nullable table) =>
        let
        // if previous is null, then this is our first page of data
        nextPage =  if contentType <> "ActivityComments" then        
                    if (previous = null) then 1 else Value.Metadata(previous)[NextPage] else null,
        // Get the last row of the table and the (next) cursor value:
        previousCursor = if contentType = "ActivityComments" and (previous <> null) then Table.Last((previous))[cursor] else null,
        lastCursor =  if contentType = "ActivityComments" then 
                        if (previous = null)  then null  else previousCursor        else null,

       // if NextPage was set to null by the previous call, we know we have no more data
        page =  if contentType <> "ActivityComments" 
                then    
                            if (nextPage <> null) 
                            then StravaData.Page(url, nextPage, contentType, StartDate, EndDate)  
                            else null    
                else 
                            if contentType = "ActivityComments" 
                            then StravaData.Page(url, nextPage, contentType, StartDate, EndDate,lastCursor)   
                            else null

        in
            page), // Returns a page record
        Pages = Table.GenerateByPage(getNextPage)  // getNextPage is a function and is being passed as an paraemter into Table.GenerateByPage()
in
    Pages; // Returns a page record




// Read a single page of data.
// If we get back an empty list, it means we are out of pages and should return null.
StravaData.Page = (url as text, optional page as number, optional contentType as text, optional StartDate as number , optional EndDate as number, optional cursor as text) 
    as nullable table =>


    let        
        queryOptions = if contentType = "Activities" then         
            [
                page = Number.ToText(page),
                per_page = Text.From(per_page)
            ]
        else if contentType = "ClubActivities"  then  
            [
                page = Number.ToText(page),
                per_page = Text.From(per_page)
            ]                
        else if contentType = "ActivityComments" and cursor <> null then  
            [
                after_cursor = cursor,
                page_size = Text.From(per_page)
            ]           
        else if contentType = "ActivityComments" and cursor = null then  
            [
                page_size = Text.From(per_page)
            ]     
        else   
            null,


        //ValueWaitFor(producer as function, interval as function, optional count as number) 
        // https://learn.microsoft.com/en-us/power-query/wait-retry
         waitForResult = ValueWaitFor(
        (iteration) =>                                                        // 1st argument in ValueWaitFor() The Task to be (possibly) retried.
            let
                    result = if contentType = "Activities" or contentType = "ActivityComments" or contentType = "ClubActivities" then
                     
                     Web.Contents(url, [
                                      //  Headers=[Authorization="Bearer " & BearerToken]  // Remove line when creating visual studio custom connector
                                        Query = queryOptions
                                        , ManualStatusHandling = {429}
                                        , IsRetry = iteration > 0 
                                        ])
                 else
                     Web.Contents(url, [
                                       // Headers=[Authorization="Bearer " & BearerToken] // Remove line when creating visual studio custom connector
                                        ManualStatusHandling = {429}
                                        , IsRetry = iteration > 0                                         
                                        ]),

                status = Value.Metadata(result)[Response.Status],

                 actualResult = if status = 429 then null else result               
            in
                actualResult,                                                                        // Will return null if a retry is required (429 error) or 
            

            (iteration) => if iteration = 0 then #duration(0, 0, 0, 0) else #duration(0, 0, 0, Number.Power(2, iteration)) , // 2nd argument in ValueWaitFor()  delay - 
        10)                                                                                          // 3rd argument in ValueWaitFor() Couunt - Maximum number of retries
            

    in

         if waitForResult = null then
         error "Value.WaitFor() Failed after multiple retry attempts"
     else
          
        let 

        json = Json.Document(waitForResult),

             result =  
            // List of records
            if (json is list) then
                if (List.IsEmpty(json)) then 
                    null
                else if contentType = "Activities"  then 
                        Table.StravaData.Activities(json) meta [NextPage = page + 1]  
                
               else if contentType = "ClubActivities" then 
                        StravaData.ClubActivities(json) meta [NextPage = page + 1]    
     
                else if contentType = "Clubs" then 
                        StravaData.Clubs(json) meta [NextPage =null] 
                
                else if contentType = "DetailedActivity" then 
                        StravaData.DetailedActivity(json) meta [NextPage =null] 

                else if contentType = "ActivityComments" then  // we need to GET the last curstor from the reord here
                        StravaData.ActivityComments(json)  // do not set metatype as irrelevant for this type                 

                else if contentType = "Athlete" then 
                        StravaData.Athlete(json) meta [NextPage =null] 

                else
                        null
            
            // single record
            else 
               Table.FromRecords({json}) meta [NextPage = null] // turn single record into one item list
         
         in 
            result;


TableStravaDataActivities = (json) =>
    let
   
    ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),     
      TableActivities = Table.FromRecords(ConvertToTable[Column1],null,MissingField.UseNull)
in
  TableActivities;






getComments = (TableIds as table) as table =>
let
    // Step 1: Add the "Comments" column by fetching comments for each activity ID
    ResultTable = Table.AddColumn(TableIds, "Comments", each 
        let
            id = Number.ToText([id]),
            activity_comment_uri = "activities/" & id & "/comments",
            result = StravaData.GetObject(activity_comment_uri, "ActivityComments")
        in
            result
    ),
    
    // Step 2: Select and expand the relevant columns from the "Comments" data
    Table2 = Table.SelectColumns(ResultTable, {"id", "Comments"}),
    ExpandedComments = Table.ExpandTableColumn(Table2, "Comments", {"athlete", "activity_id", "text", "created_at", "cursor"}, {"athlete", "activity_id", "text", "created_at", "cursor"}),
    
    // Step 3: Expand the "athlete" record and combine first and last names
    ExpandedAthlete = Table.ExpandRecordColumn(ExpandedComments, "athlete", {"firstname", "lastname"}, {"firstname", "lastname"}),
    MergedColumns = Table.CombineColumns(ExpandedAthlete, {"firstname", "lastname"}, Combiner.CombineTextByDelimiter(" ", QuoteStyle.None), "Commentor Name"),
    
    // Step 4: Sort by "created_at" and change the type of the "created_at" column
    SortedRows = Table.Sort(MergedColumns, {{"created_at", Order.Ascending}}),
    ChangedType = Table.TransformColumnTypes(SortedRows, {{"created_at", type datetime}}),
    
    // Step 5: Remove unnecessary columns
    FinalTable = Table.RemoveColumns(ChangedType, {"cursor"})
in
    FinalTable;








getDetailedActivity = (TableIds as table) as table =>
let
    // Step 1: Add the "DetailedActivity" column by fetching detailed activity data for each activity ID
    ResultTable = Table.AddColumn(TableIds, "DetailedActivity", each 
        let
            id = Number.ToText([id]),
            detailed_activity_uri = "activities/" & id,
            result = StravaData.GetObject(detailed_activity_uri, "DetailedActivity")
        in
            result
    ),
    
    // Step 3: Select the "DetailedActivity" column and expand it
    TableWithSelectedColumns = Table.SelectColumns(ResultTable, {"DetailedActivity"}),
    ColNames = Table.ColumnNames(Table.Combine(TableWithSelectedColumns[DetailedActivity])),
    ExpandedTable = Table.ExpandTableColumn(TableWithSelectedColumns, "DetailedActivity", ColNames)
in
    ExpandedTable;




getClubActivitires = (id as number) as table =>
    let
        id = Number.ToText(id),
        club_activity_uri = "clubs/" & id &  "/activities",
        result = StravaData.GetObject(club_activity_uri,"ClubActivities") 
        in
            result;
 vclubTable = StravaData.GetObject(clubs_uri,"Clubs");

 
//Functions to parse Strava data
//Parse the Detailed Activity by json
 StravaData.DetailedActivity = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error)
in
   ConvertToTable;

//Parse the clubs by json
 StravaData.Clubs = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        #"Expanded Column1" = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"id", "resource_state", "name", "profile_medium", "profile", "cover_photo", "cover_photo_small", "sport_type", "city", "state", "country", "private", "member_count", "featured", "verified", "url"}, {"id", "resource_state", "name", "profile_medium", "profile", "cover_photo", "cover_photo_small", "sport_type", "city", "state", "country", "private", "member_count", "featured", "verified", "url"})

in
   #"Expanded Column1";



//Parse the Activity Comments by json
 StravaData.ActivityComments = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        ConvertToTable2 = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"athlete","activity_id", "text", "created_at","cursor"}, {"athlete","activity_id", "text", "created_at","cursor"})

in
   ConvertToTable2;

//Parse the Club activities by json
 StravaData.ClubActivities = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"resource_state", "athlete", "name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "workout_type"}, {"resource_state", "athlete", "name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "workout_type"}),
    #"Expanded athlete" = Table.ExpandRecordColumn(#"Expanded Column1", "athlete", {"firstname", "lastname"}, {"firstname", "lastname"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded athlete",{{"distance", type number}, {"moving_time", Int64.Type}, {"elapsed_time", Int64.Type}, {"total_elevation_gain", type number}})


in
//    ConvertToTable;
   #"Changed Type";


//Parse the Athlete by json
StravaData.Athlete = (json) =>
    let
         ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        tableresult = Record.ToTable(ConvertToTable),
        result = Table.Pivot(tableresult, List.Distinct(tableresult[Name]), "Name", "Value")
in
   result;




// This is a modified version of Table.ToNavigationTable to expose the function invocation dialog in the navigator.
// This works around the issue described here: https://github.com/Microsoft/DataConnectors/issues/30
// The functions argument is a list of records with Name and Data fields.
AddNavTableFunctions = (functions as list) as table =>
    let
        asTable = Table.FromRecords(functions),
        withItemKind = Table.AddColumn(asTable, "ItemKind", each "Function"),
        withIsLeaf = Table.AddColumn(withItemKind, "IsLeaf", each true),

        tableType = Value.Type(withIsLeaf),
        newTableType = Type.AddTableKey(tableType, {"Name"}, true) meta 
        [
            NavigationTable.NameColumn = "Name", 
            NavigationTable.DataColumn = "Data",
            NavigationTable.ItemKindColumn = "ItemKind", 
            NavigationTable.IsLeafColumn = "IsLeaf"
        ],
        navigationTable = Value.ReplaceType(withIsLeaf, newTableType)
    in
        navigationTable;




// 
// Load common library functions
// 
// TEMPORARY WORKAROUND until we're able to reference other M modules
Extension.LoadFunction = (name as text) =>
    let
        binary = Extension.Contents(name),
        asText = Text.FromBinary(binary)
    in
        Expression.Evaluate(asText, #shared);


ValueWaitFor = Extension.LoadFunction("ValueWaitFor.pqm");
Table.ChangeType = Extension.LoadFunction("Table.ChangeType.pqm");
Table.GenerateByPage = Extension.LoadFunction("Table.GenerateByPage.pqm");
Table.ToNavigationTable = Extension.LoadFunction("Table.ToNavigationTable.pqm");
Table.StravaData.Activities = Extension.LoadFunction("Table.Strava.Activities.pqm");


// Diagnostics module contains multiple functions. We can take the ones we need.
Diagnostics = Extension.LoadFunction("Diagnostics.pqm");
Diagnostics.LogValue = Diagnostics[LogValue];  // Do not use Diagnostics.LogValue() in any of the auth function or will produce an error
Diagnostics.LogFailure = Diagnostics[LogFailure];


