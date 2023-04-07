// Use this file to write queries to test your data connector
let


     startdate   = #datetime(2023,3,5,0,0,0),   // Activity Start Date 
     enddate     = null ,   // Activity End Date 


        
    result = Strava.Contents(startdate,enddate),

    
    result2 = result{[Name="Activities Comments"]}[Data],
    result3 = result{[Name="Activities (Summary)"]}[Data],
    
    result4 = result{[Name="Activities (Detailed)"]}[Data],    // This has one table with two cola id and DetailedActivity [Tables]
    ColNames = Table.ColumnNames(Table.Combine(result4[DetailedActivity])),
    result6 = Table.ExpandTableColumn ( result4, "DetailedActivity", ColNames )
    
    
      
    
    
    //TableActivities = Table.FromRecords(result3[Column1])  // This has 
    //ColNames = Table.ColumnNames(TableActivities)
    

in   
    result


