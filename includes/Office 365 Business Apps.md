# Foundations of Office 365 #
# Microsoft Teams and more #

## What is Teams? ##

Teams is the next-generation team collaboration tool in the Office 365 suite. By combining new and existing apps and services, users can find all relevant team content in a single application.

Out of the box it includes:
* Audio and video conferencing 
* Permanent chat
* Online meetings and broadcasts
* Shared files in SharePoint
* Team notes in OneNote
* Team tasks in Planner
* Shared email inbox, calendar and email address
* Available for web, desktop and mobile


### Extending Teams ###

In addition to the built-in components, Teams has several ways to incorporate external content.

## Tabs ## 
### picture of Tabs ###

* Planner
* Power BI
* Trello
* Your custom Tab


## Bots ##
* Who bot
* Your custom Bot

## Connectors ##
* Twitter
* JIRA
* Salesforce
* Your custom connector


## Messaging extensions ##
* Youtube
* Your custom extension

## Teams usage ##
**Interactive:** Create a Team

**Requirements:** Microsoft Office 365 Tenant configured for Teams

**IMPORTANT:** Be sure to enable external apps for Teams and sideloading as indicated in the guide for later exercises.

![Configure tenant for Teams](media/configure-tenant-teams.png)

**Note:** this tutorial is run from a web browser but the desktop experience is nearly identical.

1. Browse to https://teams.microsoft.com and login with your Office 365 account.

	![Teams intro 1](file://media/teams-intro-1.png)
	
2. Click on “Join or create a team” in the bottom left corner
![Teams intro 2](file://media/teams-intro-2.png)
3. Enter the team name as “Lunch”, description and access policy.
![Teams intro 3](file://media/teams-intro-3.png)  
**Note:** Choose the name carefully as this will become the team’s globally visible email address.
![Teams intro 4](file://media/teams-intro-4.png)
4. Skip the invitation screen.
5. Congratulations on creating your first Team!
6. In your newly created “Lunch” team, write your first post on the General channel.
7. Each Team has a default General channel but adding more channels is a good way to focus content.  Create a “Cafeteria” channel in the Lunch team.
![Teams intro 5](file://media/teams-intro-5.png)
8. In the Cafeteria channel, click on Files and Upload a file.
![Teams intro 6](file://media/teams-intro-6.png)
9. Click on the ellipsis menu (…) to the right of the file name and choose “Make this a tab”.
![Teams intro 7](file://media/teams-intro-7.png)
10. Click on the + button in the tab section to add another tab.
![Teams intro 8](file://media/teams-intro-8.png)

## Teams connected with Office 365 components ##
In addition to the built-in components, Teams has several ways to incorporate external content:

### Tabs – appear in Teams windows ###
- Planner
- Power BI
- Trello
- Your custom Tab
### Bots ###
- Who bot
- Your custom Bot
### Connectors ### 
- Twitter
- JIRA
- Salesforce
- Your custom connector
### Messaging extensions ###
- Youtube
- Your custom extension



# Graph APIs #
## Microsoft Graph API use and purpose ##
Microsoft Graph is a central API in Office 365 that provides unified access to user resource, relationships and intelligence via a single REST endpoint.

## Security and permissions ##
- All results are scoped to the context of the current user.
- All applications using Graph API must be granted access by the user or tenant administrator.

## Tour of Graph API calls ##
| Operation        | URL           |
| ------------------------ |:-------------:| 
| GET my profile	      | https://graph.microsoft.com/v1.0/me |
| GET my files	      | https://graph.microsoft.com/v1.0/me/drive/root/children |
| GET my photo	      | https://graph.microsoft.com/v1.0/me/photo/$value |
| GET my calendar events	      | https://graph.microsoft.com/v1.0/me/events |
| GET my manager	      | https://graph.microsoft.com/v1.0/me/manager |
| GET users in my organization	      | https://graph.microsoft.com/v1.0/users |
| GET people related to me	      | https://graph.microsoft.com/v1.0/me/people |
| GET items trending around me	      | https://graph.microsoft.com/beta/me/insights/trending |
| GET my tasks assigned to me across plans	      | https://graph.microsoft.com/v1.0/me/planner/tasks/ |
| POST Create a new task	      | https://graph.microsoft.com/v1.0/planner/tasks |
| GET data from my Excel file	      | https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/ |
| POST new row to my Excel file	      | https://graph.microsoft.com/v1.0/me/drive/root:/demo.xlsx:/workbook/tables/Table1/rows/add |

## Interactive: Graph Explorer ##
1.	In a web browser, navigate to https://developer.microsoft.com/en-us/graph/graph-explorer
2.	Click on the “Run Query” button
3.	A sample JSON Response will be displayed below.
4.	Click on the button “Sign in with Microsoft” to log in to your Office 365 tenant from the previous exercise.  Take note of the permissions requested in the consent screen.
5.	Click on “show more samples” and activate the “Groups” sample category.
6.	Click on the “GET all groups I belong to” link.  You should receive a permissions error.
7.	This is Graph API’s security protecting the user from unauthorized requests.  During the initial consent screen, no “Group” permissions were requested or granted so the query about Groups is denied.
8.	Click on “modify permissions” and scroll down to activate the “Directory.Read.All”  permission. Click “Modify Permissions” to save.
9.	The session is logged out and a new consent screen is presented with the new permission request.
10.	Upon returning to the Graph Explorer screen, select “all groups I belong to” and the last JSON item will show the Team recently created.
 

## Canvases across Office 365 ##
### PowerApps canvas apps ###
Canvas apps in PowerApps are built without the need to code, using drag-and-drop tools with actions.

**Integration** - PowerApps can be integrated into many systems including:
- Teams
- SharePoint Online
- Dynamics 365
- Stand-alone

### Opening a sample PowerApp ###

### Apps and surfacing content within Office 365
### Adaptive Cards
### Store vs Enterprise apps
### Interactive: Install Office 365 Store App

## Build tailored team hubs in Microsoft Teams ##

 
 
## Fundamentals of Teams Development ##
As mentioned earlier in this lab, each Team created has a General channel and optional additional channels.
Each channel can be configured with different components to better fit its purpose:
•	Tabs
•	Bots
•	Connectors
•	Message extensions 
Tabs
Tabs allow external web pages to be embedded in the Teams user interface, allowing complete control over the user interaction experience.  Tabs are isolated from each other in IFrames and Teams provides user context to the external web page.
  

Bots
Bots allow conversational or command style interaction via a chat interface between the user and your Bot application.  Any bot using the Microsoft Bot Framework is compatible with Teams and can be enhanced for Teams using cards, channel metadata for activities and understanding messaging extensions.
Users can interact in two manners: direct 1:1 chats or mention a bot with a hashtag in a Team channel.
 

Cards
Cards are a condensed graphic representation of objects throughout Office 365, including Teams.  Each card can have properties, attachments and action buttons.  In Teams, card can be shown in chats with bots, messaging extensions and connectors.
There are several types of cards to employ depending on the object type and user interaction required.
 
Messaging Extensions
Messaging extensions enable users to search external services for objects and select an object for embedding in a channel conversation.  This mechanism combines the Microsoft Bot Framework with Cards to add rich content to chats.

 

Office 365 Connectors
Connectors bring content and events from external systems into Team conversations using Cards and incoming webhooks.  The cards generated can be static or actionable, allowing complete interaction from within Teams.

1.	Add the App Studio application to Teams
2.	Choose your custom app type:
 
Simple example with Tabs 
Adapt an existing web app to work with Microsoft Teams
In this example, we will demonstrate how to transform an existing web app into a Team App.  Most of the initial effort is related to provisioning the Team App and the original web app will remain largely untouched.
Our existing app shows the daily lunch menu, complete with images and nutritional information using a simple React app.
 

Requirements for converting your web app into a Teams tab:
1.	The content must be served via HTTPS.
2.	The page should render properly within an IFrame.
3.	The Microsoft Teams JavaScript client SDK must be included and call microsoftTeams.initialize().
4.	The app must include an initial configuration page to display during tab setup.
Include the Teams JS SDK
All pages viewed in the Teams tab must include the Teams JS SDK and call microsoftTeams.initialize().

Node.js, React and other npm-based frameworks can install using:
	npm install @microsoft/teams-js

Simple HTML pages can include:
<script src="https://statics.teams.microsoft.com/sdk/v1.2/js/MicrosoftTeams.min.js" ></script>

Initialize the Teams JS SDK
For all pages, after the page and library load is complete, call 
	microsoftTeams.initialize();

import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';            

class ConfigurePage extends Component {
    componentDidMount() {
        microsoftTeams.initialize();
    
        microsoftTeams.settings.setSettings({
            contentUrl: "https://lunchmenuhol365.blob.core.windows.net/app/index.html", // Mandatory parameter
            entityId: "7e9949f5-659e-48eb-923c-ac70e841ab09" // Mandatory parameter
      });
      microsoftTeams.settings.setValidityState(true);
        microsoftTeams.settings.registerOnSaveHandler(function(saveEvent){           
            saveEvent.notifySuccess();
        });      
    }

    render(props) {
        return (
          <div className="ConfigurePage">           
                <p>No configuration options at the moment.</p>                
          </div>
        );
      }
       
  }
  export default ConfigurePage;

Create a Configuration page or view
The configuration page is shown to the team user or administrator when a new tab is added to a channel.  Any options needed for configuring the tab may be set at this time.
Required settings:
	contentUrl – the absolute URL for the tab to be shown including query parameters
	entityId – a unique ID developers can use to identify this tab and configuration
Optional settings:
	suggestedDisplayName – the suggested tab name
         websiteUrl – the URL linked in the button “Go to website”
         removeUrl – the URL shown upon removal of a tab

A “saveHandler” is registered for when the user clicks on the Save button (this exists outside the IFrame) . Once the user has completed all the required information, the page can unlock the Save button for the user with this call:
microsoftTeams.initialize();
    
        microsoftTeams.settings.setSettings({
            contentUrl: "https://lunchmenuhol365.blob.core.windows.net/app/index.html", // Mandatory parameter
            entityId: "7e9949f5-659e-48eb-923c-ac70e841ab09" // Mandatory parameter
      });
      microsoftTeams.settings.setValidityState(true);
        microsoftTeams.settings.registerOnSaveHandler(function(saveEvent){           
            saveEvent.notifySuccess();
        });      

Add Team UI components
To maintain the Teams and Office 365 look and feel, 
Deploy the webapp

1.	Deploy your app to a publicly accessible platform like Azure App Service, Azure Storage, Amazon AWS or your own web server.  
Note: The content must be hosted via HTTPS.
 
Create an app manifest using Teams App Studio
•	Interactive: Hello World! in Teams
1.	Create an existing React web application published in Azure
2.	Create Azure Resource Group
 


Additional Reading topics 
Bots
Connectors
Custom Excel visuals example
O365 canvas example

Further resources
•	Tabs dev
•	Teams Dev
•	Bots Dev
•	Connectors Dev
