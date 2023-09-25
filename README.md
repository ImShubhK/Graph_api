Graph API for retrieving the metadata and body of Outlook and Microsoft calendar



In this project, I am retrieving the data from Outlook and the calendar of the user. So here are the following steps by which I achieved  that

Authenticate the user using MSAL.

I authenticate the user using MSAL I implemented login and logout functionality for the user in that process I stored the access token of the user so that I can retrieve their data.



Metadata and event body of the  Calendar:

I hit this endpoint to retrieve the metadata and event body of the calendar. https://graph.microsoft.com/v1.0/me/events





Metadata and event body of the  of the Outlook:

For achieving this milestone I hit this endpoints 'https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,body,sender,toRecipients'







Here you can see all those results in the console as well.



For retrieving all these data I used the React.js framework.

#For run this application 
1 clone the repo and do npm I
2. After that just do npm start.

