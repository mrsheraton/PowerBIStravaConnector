// Use this file to write queries to test your data connector
let


     startdate   = #datetime(2024,08,01,0,0,0),   // Activity Start Date 
     enddate     = null ,   // Activity End Date 

        
    Source = Strava.Contents(startdate,enddate),
    Activities_Summary = Source{[Name="Activities (Summary)"]}[Data],
    Activities_Comments = Source{[Name="Activities (Comments)"]}[Data],
    Activities_Detailed = Source{[Name="Activities (Detailed)"]}[Data],
    Clubs = Source{[Name="Clubs"]}[Data]
    /*
    Activities (Summary)
    Activities (Comments)
    Activities (Detailed)
    Activities (Detailed) (Laps)
    Activities (Detailed) (Segment Efforts)
    Activities (Detailed) (Best Efforts)
    Clubs
    Athlete
    */

in   
    Clubs


