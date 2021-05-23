section Strava;

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

activity_uri = "athlete/activities";                            //  returns a SummaryActivity
athlete_uri = "athlete";
clubs_uri = "/athlete/clubs";


// default page size, using the MAX page size to limit the amount of calls to the API
page_size = 200;
windowWidth = 1200;
windowHeight = 1000;


// Exported functions
//
// These functions are exported to the M Engine (making them visible to end users), and associates 
// them with the specified Data Source Kind. The Data Source Kind is used when determining which 
// credentials to use during evaluation. Credential matching is done based on the function's parameters. 
// All data source functions associated to the same Data Source Kind must have a matching set of required 
// function parameters, including type, name, and the order in which they appear. 

[DataSource.Kind="Strava", Publish="Strava.UI"]
shared Strava.Contents = StravaData.NavTable;  


GetCurrent_credentials = () as any =>
    let
        Properties = Extension.CurrentCredential()
       
        in
            Properties;
        


//
// Data Source definition
//
Strava = [
    TestConnection = (dataSourcePath) => { "Strava.Contents" },
    Authentication = [
        OAuth = [                                                               
            StartLogin = StartLogin,
            FinishLogin = FinishLogin,                                                     
            Refresh = Refresh,                                 

            Label = Extension.LoadString("AuthenticationLabel")
        ]
    ],
    Label = Extension.LoadString("DataSourceLabel")
];


StravaAnon = [
    TestConnection = (dataSourcePath) => { "Strava.Contents" },
    Authentication = [
        Anonymous = []
        ],
    Label = Extension.LoadString("DataSourceLabel")
];



StartLogin = (resourceUrl, state, display) =>
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


FinishLogin = (context, callbackUri, state) =>  
    let
        Parts = Uri.Parts(callbackUri)[Query]                                                                    // Step 1 - request access  (Parts is a JSON file)
    in
        TokenMethod(Parts[code]);                                                                                   // code is the short live authorisaiton code




TokenMethod = (code) =>
    let
        Response = Web.Contents("https://www.strava.com/oauth/token", [                                         // Step 2 - Token Exchange (Get access token)
            Content = Text.ToBinary(Uri.BuildQueryString([
                client_id = client_id,
                client_secret = client_secret,
                code = code,                            
                //redirect_uri = redirect_uri               // MAS removed as not required
                grant_type = "authorization_code"            // MAS added this as missing
                ])),
            Headers=[#"Content-type" = "application/x-www-form-urlencoded",#"Accept" = "application/json"]]), 
        Parts = Json.Document(Response)
    in
        Parts;                                                                                                   


Refresh = (code,refresh_token) =>      
    let
        Response = Web.Contents("https://www.strava.com/oauth/token", [                                    
            Content = Text.ToBinary(Uri.BuildQueryString([
                client_id = client_id,
                client_secret = client_secret,
                 code = code,                            
                 grant_type = "refresh_token",
                 refresh_token = refresh_token])),
            Headers=[#"Content-type" = "application/x-www-form-urlencoded",#"Accept" = "application/json"]]), 
        Parts = Json.Document(Response)
    in
        Parts;             



//
// UI Export definition
//
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

//For the navigation tables see this example: https://github.com/Microsoft/DataConnectors/tree/master/samples/TripPin/3-NavTables 

//
// Nav table definition
//
// Define the NAV URL
StravaData.NavTable = (optional  StartDate as nullable datetime ,optional EndDate as nullable datetime) as table =>
    let

//// Use me for testing in Visual Studio only
    //// 1 (both blank, will give error for me as over 600 API requests needed)
         //StartDate = null,
         //EndDate = null,

// //2 (Start date only)
//        StartDate = #datetime( 2021,4,1,0,0,0),
//        EndDate = null,



//// 3 (End date only)
//          StartDate = null,
//          EndDate = #datetime( 2015,12,31,0,0,0),


         StartDate2 = if StartDate = null then #datetime( 1980,1,1,0,0,0) else StartDate,
         EndDate2 = if EndDate = null then #datetime( 3000,1,1,0,0,0) else EndDate,
                             
        vclub = StravaData.GetObject(clubs_uri,"Clubs"),
            vclubWithActivities = Table.AddColumn(vclub, "Activities", each getClubActivitires([id])),  

        vMyActivities = StravaData.GetObject(activity_uri,"Activities",StartDate2,EndDate2 ),
            vMyActivitiesWithComments = Table.AddColumn(vMyActivities, "Comments", each if  [comment_count] > 1 then getActivityComment([id]) else null  ),  

        source = #table({"Name"         , "Data"}, {
                        { "Athlete"     , StravaData.GetObject(athlete_uri,"Athlete")},
                        { "Activities"  , vMyActivitiesWithComments },                                      
                        { "Clubs"       , vclubWithActivities}


                        
        }),
        
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
        
        navTable = Table.ToNavigationTable(withFunction, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")

        
        
        in
                navTable;


//All of the paging is based off the Tripin Sample https://github.com/Microsoft/DataConnectors/tree/master/samples/TripPin/5-Paging 

// Add a helper function that can be used to retrieve strava API objects 
// that we haven't exposed in the navigation table. We always take a relative path,
// and append it to the base Strava API URL.
// Note: Taking in a full path is a security issue, as it could potentially leak 
// the bearer token to the url specified in the function parameter. 
// For Strava I added an additional parameter to determine the type of content so we can parse this differently later
StravaData.GetObject = (relativeUrl as text, contentType as text, optional StartDate as datetime , optional EndDate as datetime) => StravaData.PagedReader(api_uri & relativeUrl, contentType, StartDate, EndDate);


// Read all pages of data.
// After every page, we check the "NextPage" record on the metadata of the previous request.
// Table.GenerateByPage will keep asking for more pages until we return null.
StravaData.PagedReader = (url as text, contentType as text, optional StartDate as datetime , optional EndDate as datetime) =>
    Table.GenerateByPage((previous) => 
        let
            // if previous is null, then this is our first page of data
            nextPage = if (previous = null) then 1 else Value.Metadata(previous)[NextPage]?,
            // if NextPage was set to null by the previous call, we know we have no more data
            page = if (nextPage <> null) then StravaData.Page(url, nextPage, contentType, StartDate, EndDate) else null
        in
            page
    );


// Read a single page of data.
// If we get back an empty list, it means we are out of pages and should return null.
// For Strava I added an additional parameter to determine the type of content so we can parse this differently later
StravaData.Page = (url as text, page as number, contentType as text, optional StartDate as datetime , optional EndDate as datetime) as nullable table =>
    let

        // converts text date to a number (Epoch date)
        EndDate = Duration.TotalSeconds( EndDate - #datetime(1970,1,1,0,0,0)),         
        StartDate = Duration.TotalSeconds(  StartDate - #datetime(1970,1,1,0,0,0)),          

        queryOptions =
        if contentType = "Activities" then 
        
            [
            page = Number.ToText(page),
            page_size = Text.From(page_size),
            before = Text.From(EndDate),
            after =  Text.From(StartDate)
            ]
            else 
            [
            page = Number.ToText(page),
            page_size = Text.From(page_size)
            ]
            
            ,

        content = Web.Contents(url, [ Query = queryOptions ]),
        json = Json.Document(content),

        // Responses from the API will either be a single record, or list of records.
        // When we have a single record, we want to format it as a table. Set NextPage to null.
        // When we have a list of records, we convert to a table. Set NextPage to page + 1.
        // When we have an empty list of records, we return null. 
        result =
            if (json is list) then
                if (List.IsEmpty(json)) then 
                    null
                else //Determine how to parse the JSON based on the type of content.
                    if(contentType = "Activities") then
                        StravaData.Activities(json) meta [NextPage = page + 1]

                        else

                            if(contentType = "ActivityComments") then
                            StravaData.ActivityComments(json) meta [NextPage = page + 1]
                            else

                            if(contentType = "ClubActivities") then
                            StravaData.ClubActivities(json) meta [NextPage = page + 1]

                            else
                                if(contentType = "Clubs") then
                                StravaData.Clubs(json) meta [NextPage = page + 1]
                            
                    else
                        StravaData.Athlete(json) meta [NextPage = page + 1]
            else
                Table.FromRecords({json}) meta [NextPage = null] // turn single record into one item list
    in
        result;


//Functions to parse Strava data

getClubActivitires = (id as number) as table =>
    let
        id = Number.ToText(id),
        club_activity_uri = "clubs/" & id &  "/activities",
        result = StravaData.GetObject(club_activity_uri,"ClubActivities") 
        
        in
            result;

 vclubTable = StravaData.GetObject(clubs_uri,"Clubs");

 
 getActivityComment = (id as number) as table =>
    let
        id = Number.ToText(id),


        activity_comment_uri = "activities/" & id &  "/comments",
        result =  StravaData.GetObject(activity_comment_uri,"ActivityComments") 

        in
            result;



//Parse the clubs by json
 StravaData.Clubs = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        #"Expanded Column1" = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"id", "resource_state", "name", "profile_medium", "profile", "cover_photo", "cover_photo_small", "sport_type", "city", "state", "country", "private", "member_count", "featured", "verified", "url"}, {"id", "resource_state", "name", "profile_medium", "profile", "cover_photo", "cover_photo_small", "sport_type", "city", "state", "country", "private", "member_count", "featured", "verified", "url"})

in
   #"Expanded Column1";

//Parse the activities by json
 StravaData.Activities = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        #"Expanded Column1" = Table.ExpandRecordColumn( ConvertToTable, "Column1", {"id", "resource_state", "external_id", "upload_id", "athlete", "name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "start_date", "start_date_local", "timezone", "utc_offset", "start_latlng", "end_latlng", "location_city", "location_state", "location_country", "start_latitude", "start_longitude", "achievement_count", "kudos_count", "comment_count", "athlete_count", "photo_count", "map", "trainer", "commute", "manual", "private", "flagged", "gear_id", "average_speed", "max_speed", "average_cadence", "average_temp", "average_watts", "kilojoules", "device_watts", "has_heartrate", "average_heartrate", "max_heartrate", "elev_high", "elev_low", "pr_count", "total_photo_count", "has_kudoed", "workout_type", "weighted_average_watts", "max_watts"}, {"id", "resource_state", "external_id", "upload_id", "athlete", "name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "start_date", "start_date_local", "timezone", "utc_offset", "start_latlng", "end_latlng", "location_city", "location_state", "location_country", "start_latitude", "start_longitude", "achievement_count", "kudos_count", "comment_count", "athlete_count", "photo_count", "map", "trainer", "commute", "manual", "private", "flagged", "gear_id", "average_speed", "max_speed", "average_cadence", "average_temp", "average_watts", "kilojoules", "device_watts", "has_heartrate", "average_heartrate", "max_heartrate", "elev_high", "elev_low", "pr_count", "total_photo_count", "has_kudoed", "workout_type", "weighted_average_watts", "max_watts"}),
        //#"Removed Columns" = Table.RemoveColumns(#"Expanded Column1",{"id", "resource_state", "upload_id", "athlete", "start_latlng", "end_latlng", "location_city", "location_state", "photo_count", "map", "total_photo_count", "external_id"}),
        #"Changed Type" = Table.TransformColumnTypes(#"Expanded Column1",{{"total_elevation_gain", Int64.Type}, {"elapsed_time", Int64.Type}, {"moving_time", Int64.Type}, {"distance", type number}}),
        #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"start_date", "start_datetime"}}),
        #"Duplicated Column" = Table.DuplicateColumn(#"Renamed Columns", "start_datetime", "start_datetime - Copy"),
        #"Changed Type1" = Table.TransformColumnTypes(#"Duplicated Column",{{"start_datetime - Copy", type datetime}}),
        #"Extracted Date" = Table.TransformColumns(#"Changed Type1",{{"start_datetime - Copy", DateTime.Date}}),
        #"Renamed Columns1" = Table.RenameColumns(#"Extracted Date",{{"start_datetime - Copy", "start_date"}}),
        #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns1",{"name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "start_datetime", "start_date", "start_date_local", "timezone", "utc_offset", "location_country", "start_latitude", "start_longitude", "achievement_count", "kudos_count", "comment_count", "athlete_count", "trainer", "commute", "manual", "private", "flagged", "gear_id", "average_speed", "max_speed", "average_cadence", "average_temp", "average_watts", "kilojoules", "device_watts", "has_heartrate", "average_heartrate", "max_heartrate", "elev_high", "elev_low", "pr_count", "has_kudoed", "workout_type", "weighted_average_watts", "max_watts"}),
        #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns",{"start_date_local", "timezone", "utc_offset"}),
        #"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns1",{{"start_latitude", type number}, {"start_longitude", type number}, {"achievement_count", Int64.Type}, {"kudos_count", Int64.Type}, {"comment_count", Int64.Type}, {"athlete_count", Int64.Type}, {"trainer", type logical}, {"commute", type logical}, {"manual", type logical}, {"private", type logical}, {"flagged", type logical}, {"average_speed", type number}, {"max_speed", type number}, {"average_cadence", type number}, {"average_temp", type number}, {"average_watts", type number}, {"kilojoules", type number}, {"device_watts", type logical}, {"has_heartrate", type logical}, {"average_heartrate", type number}, {"max_heartrate", Int64.Type}, {"elev_high", type number}, {"elev_low", type number}, {"pr_count", Int64.Type}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Changed Type2",{"workout_type"}),
        #"Changed Type3" = Table.TransformColumnTypes(#"Removed Columns2",{{"has_kudoed", type logical}, {"weighted_average_watts", Int64.Type}, {"max_watts", Int64.Type}}),
        #"Changed Typedt" = Table.TransformColumnTypes(#"Changed Type3",{{"start_datetime", type datetime}}),
        #"Changed Type1t" = Table.TransformColumnTypes(#"Changed Typedt",{{"start_datetime", type time}}),
        #"Duplicated Column2" = Table.RenameColumns(#"Changed Type1t", {{"start_datetime", "start_time"}}),
        #"Removed Blank Rows" = Table.SelectRows(#"Duplicated Column2", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))
in
    #"Removed Blank Rows";



//Parse the Activity Comments by json
 StravaData.ActivityComments = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error)

in
   ConvertToTable;

//Parse the Club activities by json
 StravaData.ClubActivities = (json) =>
    let
        ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"resource_state", "athlete", "name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "workout_type"}, {"resource_state", "athlete", "name", "distance", "moving_time", "elapsed_time", "total_elevation_gain", "type", "workout_type"}),
    #"Expanded athlete" = Table.ExpandRecordColumn(#"Expanded Column1", "athlete", {"firstname", "lastname"}, {"firstname", "lastname"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded athlete",{{"distance", type number}, {"moving_time", Int64.Type}, {"elapsed_time", Int64.Type}, {"total_elevation_gain", type number}})


in
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
// Helper functions
//
 Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(table),
        newTableType = Type.AddTableKey(tableType, keyColumns, true) meta 
        [
            NavigationTable.NameColumn = nameColumn, 
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn, 
            Preview.DelayColumn = itemNameColumn, 
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(table, newTableType)
    in
        navigationTable;

Table.GenerateByPage = (getNextPage as function) as table =>
    let
        listOfPages = List.Generate(
            () => getNextPage(null),
            (lastPage) => lastPage <> null,
            (lastPage) => getNextPage(lastPage)
        ),
        tableOfPages = Table.FromList(listOfPages, Splitter.SplitByNothing(), {"Column1"}),
        firstRow = tableOfPages{0}?
    in
        if (firstRow = null) then
            Table.FromRows({})
        else
            Value.ReplaceType(
                Table.ExpandTableColumn(tableOfPages, "Column1", Table.ColumnNames(firstRow[Column1])),
                Value.Type(firstRow[Column1])
            );
