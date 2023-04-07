# Strava Connector
The Strava M extension allows you to connect to Strava and retrieve your activities and athlete information for analytics.

Before you get started using the M extension, you need to register a new app on Strava, and add two files `client_id` and `client_secret` to the project and add them with the appropriate values for you app.

## OAuth and Power BI
OAuth is a form of credentials delegation. By logging in to Strava and authorizing the "application" you create for Strava, the user is allowing your "application" to login on their behalf to retrieve data into Power BI.
The "application" must be granted rights to retrieve data (get an access_token) and to refresh the data on a schedule (get and use a refresh_token).
Your "application" in this context is your Data Connector used to run queries within Power BI.
Power BI stores and manages the access_token and refresh_token on your behalf.

**Note:** To allow Power BI to obtain and use the access_token, you must specify the redirect url as: https://oauth.powerbi.com/views/oauthredirect.html

When you specify this URL and Strava successfully authenticates and grants permissions, Strava will redirect to PowerBI's oauthredirect endpoint so that Power BI can retrieve the access_token and refresh_token. 

## How to register a Strava app
Your Power BI extension needs to login to Strava. To enable this, you register a new OAuth application with Strava at https://www.strava.com/settings/api.
1. `Application name`: Enter a name for the application for your M extension.  
2. `Authorization callback Domain`: Enter oauth.powerbi.com.   
**Note:** A registered OAuth application is assigned a unique Client ID and Client Secret. The Client Secret should not be shared. You get the Client ID and Client Secret from the Strava application page.
3 Add 2 files `client_id` and `client_secret` to the project in your Data Connector project with the Client ID (`client_id` file) and Client Secret (`client_secret` file).

