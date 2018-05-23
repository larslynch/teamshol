# Foundations of Office 365 #
## Microsoft Teams and more ##

### What is Teams? ###

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

### Tabs ###

* Planner
* Power BI
* Trello
* Your custom Tab


### Bots ###
* Who bot
* Your custom Bot

### Connectors ###
* Twitter
* JIRA
* Salesforce
* Your custom connector


### Messaging extensions ###
* Youtube
* Your custom extension

## Exercise 1: Create a Team ##

### Prerequisites ###
1. [Microsoft Office 365 Tenant configured for Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-tenant)

> **IMPORTANT:** Be sure to enable external apps for Teams and sideloading as indicated in the guide for later exercises.

![Configure tenant for Teams](../media/configure-tenant-teams.png)

> **Note:** this tutorial is run from a web browser but the desktop experience is nearly identical.

1. Browse to https://teams.microsoft.com and login with your Office 365 account.

![Teams intro 1](../media/teams-intro-1.png)
	
2. Click on “Join or create a team” in the bottom left corner

![Teams intro 2](../media/teams-intro-2.png)

3. Enter the team name as “Lunch”, description and access policy.

![Teams intro 3](../media/teams-intro-3.png)  

> **Note:** Choose the name carefully as this will become the team’s globally visible email address.

![Teams intro 4](../media/teams-intro-4.png)

4. Skip the invitation screen.
5. Congratulations on creating your first Team!
6. In your newly created “Lunch” team, write your first post on the General channel.
7. Each Team has a default General channel but adding more channels is a good way to focus content.  Create a “Cafeteria” channel in the Lunch team.

![Teams intro 5](../media/teams-intro-5.png)

8. In the Cafeteria channel, click on Files and Upload a file.

![Teams intro 6](../media/teams-intro-6.png)

9. Click on the ellipsis menu (…) to the right of the file name and choose “Make this a tab”.

![Teams intro 7](../media/teams-intro-7.png)

10. Click on the + button in the tab section to add another tab.

![Teams intro 8](../media/teams-intro-8.png)

11. Now that you have seen Teams from a user's perspective, read on to see what a developer can do!

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

## Exercise 2: Graph Explorer ##
1.	In a web browser, navigate to https://developer.microsoft.com/en-us/graph/graph-explorer
2.	Click on the “Run Query” button

![Graph 1](../media/graph-1.png)

3.	A sample JSON Response will be displayed below.
4.	Click on the button “Sign in with Microsoft” to log in to your Office 365 tenant from the previous exercise.  Take note of the permissions requested in the consent screen.

![Graph 2](../media/graph-2.png)

5.	Click on “show more samples” and activate the “Groups” sample category.

![Graph 3](../media/graph-3.png)

6.	Click on the “GET all groups I belong to” link.  You should receive a permissions error.

![Graph 4](../media/graph-4.png)

7.	This is Graph API’s security protecting the user from unauthorized requests.  During the initial consent screen, no “Group” permissions were requested or granted so the query about Groups is denied.

8.	Click on “modify permissions” and scroll down to activate the “Directory.Read.All”  permission. Click “Modify Permissions” to save.

![Graph 5](../media/graph-5.png)

9.	The session is logged out and a new consent screen is presented with the new permission request.

![Graph 6](../media/graph-6.png)

10.	Upon returning to the Graph Explorer screen, select “all groups I belong to” and the last JSON item will show the Team recently created.
 
![Graph 7](../media/graph-7.png)



# Canvases across Office 365 #

### Marketplace ###
[Microsoft AppSource](https://appsource.microsoft.com/en-us/marketplace/apps?src=office&page=1&product=office)

### PowerApps ###
Canvas apps in PowerApps are built without the need to code, using drag-and-drop tools with actions.

PowerApps can be integrated into many systems including:
- Teams
- SharePoint Online
- Dynamics 365
- Stand-alone

- [Overview of creating apps in PowerApps](https://docs.microsoft.com/en-us/powerapps/maker/index)
- [Create an app from a PowerApps template](https://docs.microsoft.com/en-us/powerapps/maker/canvas-apps/get-started-test-drive)

###  Flow & Logic Apps ###
- [Get started with Microsoft Flow](https://docs.microsoft.com/en-us/flow/getting-started)
- [Create a flow in Microsoft Flow](https://docs.microsoft.com/en-us/flow/get-started-logic-flow)
- [What is Azure Logic Apps?](https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-overview)
- [Quickstart: Build your first logic app workflow - Azure portal](https://docs.microsoft.com/en-us/azure/logic-apps/quickstart-create-first-logic-app-workflow)
- [Compare Flow, Logic Apps, Functions, and WebJobs](https://docs.microsoft.com/en-us/azure/azure-functions/functions-compare-logic-apps-ms-flow-webjobs)

### Power BI ###
- [What can developers do with Power BI?](https://docs.microsoft.com/en-us/power-bi/developer/what-can-you-do)
- [Use developer tools to create custom visuals](https://docs.microsoft.com/en-us/power-bi/service-custom-visuals-getting-started-with-developer-tools)


### Yammer ###
- [Build Your First App](https://developer.yammer.com/docs) 


###  Office Add-ins ###
- [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)

### Outlook ###
- [Outlook add-ins overview]( https://docs.microsoft.com/en-us/outlook/add-ins/)
- [Build your first Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/quick-start?tabs=visual-studio)
	
### Excel ###
- [Excel add-ins overview](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview)
- [Build an Excel add-in using React](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react)
	
### Word ###
- [Word add-ins overview](https://docs.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins-programming-overview)
- [Build your first Word add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins?tabs=visual-studio)

### PowerPoint ###
- [PowerPoint add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/powerpoint/powerpoint-add-ins)
- [Build your first PowerPoint add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/powerpoint/powerpoint-add-ins-get-started?tabs=visual-studio)

### OneNote ###
- [Build your first OneNote add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-getting-started)

### Project ###
- [Task pane add-ins for Project](https://docs.microsoft.com/en-us/office/dev/add-ins/project/project-add-ins)
- [Build your first Project add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/project/project-add-ins-get-started)
	
### SharePoint Framework ###
- [Overview of the SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Overview of SharePoint client-side web parts](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts)
- [Build your first SharePoint client-side web part](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part)
- [Overview of SharePoint Framework Extensions](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/overview-extensions)
- [Build your first SharePoint Framework Extension](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/build-a-hello-world-extension)
- [SharePoint Add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/sharepoint-add-ins)
- [Get started creating provider-hosted SharePoint Add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-started-creating-provider-hosted-sharepoint-add-ins)
- [SharePoint Framework (SPFx) enterprise guidance](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/enterprise-guidance)

### Teams apps ###
- [Apps for Microsoft Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-overview)

### Office 365 APIs ###
- [Overview of Office 365 file handlers](https://msdn.microsoft.com/en-us/office/office365/howto/using-cross-suite-apps)
- [Create file handlers in Office 365](https://msdn.microsoft.com/en-us/office/office365/howto/create-file-handler-extensions)

### Graph APIs ###



# Build tailored team hubs in Microsoft Teams #

## Fundamentals of Teams Development ##
As mentioned earlier in this lab, each Team created has a General channel and optional additional channels.
Each channel can be configured with different components to better fit its purpose:
- Tabs
- Bots
- Connectors
- Message extensions

### Tabs ###
Tabs allow external web pages to be embedded in the Teams user interface, allowing complete control over the user interaction experience.  Tabs are isolated from each other in IFrames and Teams provides user context to the external web page.
  
![Tabs 1 ](../media/tabs-example-1.png)


### Bots ###
Bots allow conversational or command style interaction via a chat interface between the user and your Bot application.  Any bot using the Microsoft Bot Framework is compatible with Teams and can be enhanced for Teams using cards, channel metadata for activities and understanding messaging extensions.

Users can interact in two manners: direct 1:1 chats or mention a bot with a hashtag in a Team channel.
  
![Bots 1](../media/bots-example-1.png)


### Cards ### 
Cards are a condensed graphic representation of objects throughout Office 365, including Teams.  Each card can have properties, attachments and action buttons.  In Teams, card can be shown in chats with bots, messaging extensions and connectors.
There are several types of cards to employ depending on the object type and user interaction required.
  
![Cards 1](../media/cards-example-1.png)


### Messaging Extensions ### 
Messaging extensions enable users to search external services for objects and select an object for embedding in a channel conversation.  This mechanism combines the Microsoft Bot Framework with Cards to add rich content to chats.
  
![Messaging Extensions 1](../media/messaging-extensions-example-1.png)


### Office 365 Connectors ###
Connectors bring content and events from external systems into Team conversations using Cards and incoming webhooks.  The cards generated can be static or actionable, allowing complete interaction from within Teams.
  
![Office 365 Connectors 1](../media/office365-connector-example-1.png)



 
## Exercise 3: Simple example with Tabs ##
###  Adapt an existing web app to work with Microsoft Teams ###
In this example, we will demonstrate how to transform an existing web app into a Team App.  

Our existing app shows the daily lunch menu, complete with images and nutritional information using a simple React app.

![Lunch menu  1](../media/lunchmenu-1.png)

### Build and deploy the existing web app ###

#### Prerequisites ####
1. [Microsoft Office 365 Tenant configured for Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-tenant)
2. Azure subscription
3. [Node LTS and npm](https://nodejs.org/en/)
4. [Visual Studio Code](https://code.visualstudio.com/)


**Requirements for converting your web app into a Teams tab:**
1. The content must be served via HTTPS.
2. The page should render properly within an IFrame.
3. The Microsoft Teams JavaScript client SDK must be included and call `microsoftTeams.initialize()`.
4. The app must include an initial configuration page to display during tab setup.

## Create Azure App Service ##
1. Login to Azure portal

![Azure Portal](../media/azure-portal-1.png)

2. Select **App Services**, followed by **+Add** then choose the **Web App** template.  Lastly, click **Create**.

![Azure Portal 2](../media/azure-portal-2.png)

3. Complete the web app details and click on **Service Plan** to choose a new free service plan.
4. Enter a name for the new service plan and change the pricing tier to **F1 Free**

![Azure Portal 4](../media/azure-portal-4.png)

5. Click **OK** to create the service plan and **Create** to provision the web app.

![Azure Portal 5](../media/azure-portal-5.png)

## Create simple React web app ##
1. Run **Visual Studio Code** and open an **Integrated Terminal**

![Visual Studio Code](../media/vscode-terminal-1.png)

2. In the Powershell terminal that opens, run this command to globally install the NPM React template:
```
	npm install -g create-react-app
```	
3. Change to your development directory, create the React app and run it:
```
	cd c:\temp
	create-react-app lunchmenu 
	cd lunchmenu 
	npm start
```
4. A web browser opens to http://localhost:3000 and displays the web app

![React 1](../media/react-1.png)

5. Return to Visual Studio Code and open the **lunchmenu** folder

![Visual Studio Code 2](../media/vscode-folder-1.png)

6. Download this [assets.zip](../media/assets.zip) file and unpack the images and JSON to the **lunchmenu\src\assets** folder.
7. Create these new files in Visual Studio Code:
- **src\MenuContainer.js** - A React parent container component
- **src\MenuItem.js** - A React component for rendering each item on the lunch menu
- **src\MenuItem.css** - The CSS for each menu item
- **public\web.config** - A modified web.config to route all web requests to index.html, necessary for SPA

![Visual Studio Code File](../media/vscode-file-1.png)

8.  Complete the new files with the following:
**src\MenuItem.js**
```js
import React, { Component } from 'react';

import './MenuItem.css';

class MenuItem extends Component {

    constructor(props) {
        super(props);
        this.state = {showImage: true};
    
        // This binding is necessary to make `this` work in the callback
        this.handleClick = this.handleClick.bind(this);
      }

      handleClick() {
        this.setState(prevState => ({
            showImage: !prevState.showImage
        }));
      }

    render(props) {
      return (
        <div className="MenuItem" onClick={this.handleClick}>
        <span className={this.state.showImage ? 'visible' : 'hidden'}><h1>{this.props.data_url.title} {this.props.data_url.price.toLocaleString(navigator.language,{style: 'currency', currency:'USD'})}</h1><img alt={this.props.data_url.title} src={this.props.image_url} className="menu-image" /></span>
        <span className={this.state.showImage ? 'hidden' : 'visible'}><h1>{this.props.data_url.title} {this.props.data_url.price.toLocaleString(navigator.language,{style: 'currency', currency:'USD'})}</h1>          
            <div className="nutrition-container">
            <div>
            <table className="nutrition-table-1">
            <tr><td colSpan="2"><span className="nutrition-label">Serving size</span></td><td colSpan="2">{this.props.data_url.nutrition_table.serving_size}</td><td></td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Calories</span></td><td colSpan="2">{this.props.data_url.nutrition_table.calories}</td><td></td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Total fat </span>{this.props.data_url.nutrition_table.total_fat[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.total_fat[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Saturated fat </span>{this.props.data_url.nutrition_table.saturated_fat[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.saturated_fat[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Polyunsaturated fat </span>{this.props.data_url.nutrition_table.polyunsaturated_fat[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.polyunsaturated_fat[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Monounsaturated fat </span>{this.props.data_url.nutrition_table.monounsaturated_fat[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.monounsaturated_fat[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Cholesterol </span>{this.props.data_url.nutrition_table.cholesterol[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.cholesterol[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Sodium </span>{this.props.data_url.nutrition_table.sodium[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.sodium[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Potassium </span>{this.props.data_url.nutrition_table.potassium[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.potassium[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Total carbohydrate </span>{this.props.data_url.nutrition_table.total_carbohydrate[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.total_carbohydrate[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Dietary fiber </span>{this.props.data_url.nutrition_table.dietary_fiber[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.dietary_fiber[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Sugar </span>{this.props.data_url.nutrition_table.sugar[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.sugar[1]}</td></tr>
            <tr><td colSpan="2"><span className="nutrition-label">Protein </span>{this.props.data_url.nutrition_table.protein[0]}</td><td colSpan="2">{this.props.data_url.nutrition_table.protein[1]}</td></tr>
            </table>
            </div>
            
            <div>
            <table className="nutrition-table-2">
            <tr><td>Vitamin A</td><td>{this.props.data_url.nutrition_table.vitamin_a[0]}</td><td>{this.props.data_url.nutrition_table.vitamin_a[1]}</td><td>Vitamin C</td><td>{this.props.data_url.nutrition_table.vitamin_c[0]}</td><td>{this.props.data_url.nutrition_table.vitamin_c[1]}</td></tr>
            <tr><td>Calcium</td><td>{this.props.data_url.nutrition_table.calcium[0]}</td><td>{this.props.data_url.nutrition_table.calcium[1]}</td><td>Iron</td><td>{this.props.data_url.nutrition_table.iron[0]}</td><td>{this.props.data_url.nutrition_table.iron[1]}</td></tr>
            <tr><td>Vitamin D</td><td>{this.props.data_url.nutrition_table.vitamin_d[0]}</td><td>{this.props.data_url.nutrition_table.vitamin_d[1]}</td><td>Vitamin B-6</td><td>{this.props.data_url.nutrition_table.vitamin_b_6[0]}</td><td>{this.props.data_url.nutrition_table.vitamin_b_6[1]}</td></tr>
            <tr><td>Vitamin B-12</td><td>{this.props.data_url.nutrition_table.vitamin_b_12[0]}</td><td>{this.props.data_url.nutrition_table.vitamin_b_12[1]}</td><td>Magnesium</td><td>{this.props.data_url.nutrition_table.magnesium[0]}</td><td>{this.props.data_url.nutrition_table.magnesium[1]}</td></tr>
            </table>
            </div>
            </div>
        </span>
        </div>
      );
    }

     
}

export default MenuItem;
```
**src\MenuItem.css**
```css
.hidden  {
    display:none;
}

.visible {
    display: inherit;
}

.menu-image {
    width:30%;
}

.nutrition-label {
    font-weight: bold;
}


.nutrition-table-1, .nutrition-table-2 {
    flex: 0.5;
}

.nutrition-container {
    display: flex;
    justify-content: center;
    
}

.nutrition-container div {
    max-width: 30%;
}
```


**src\MenuContainer.js**
```js
import React, { Component } from 'react';

import image_1 from './assets/barbecue-chicken.jpg';
import image_2 from './assets/tortellini.jpg';
import image_3 from './assets/eggplant.jpg';
import data_1 from "./assets/chicken_data.json";
import data_2 from "./assets/tortellini_data.json";
import data_3 from "./assets/eggplant_data.json";

import MenuItem from './MenuItem';   


class MenuContainer extends Component {
      render(props) {
          return (
            <div className="MenuContainer">
                <header className="App-header">
                <h1 className="App-title">Today's menu</h1>
            </header>
                
            <MenuItem data_url={data_1} image_url={image_1}/>
            <MenuItem data_url={data_2} image_url={image_2}/>
            <MenuItem data_url={data_3} image_url={image_3}/>
            </div>
          );
 }
}
export default MenuContainer;
```



**public\web.config**
```xml
<?xml version="1.0"?>
<configuration>
 <system.webServer>
 <rewrite>
 <rules>
 <rule name="React Routes" stopProcessing="true">
 <match url=".*" />
 <conditions logicalGrouping="MatchAll">
 <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
 <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
 <add input="{REQUEST_URI}" pattern="^/(api)" negate="true" />
 </conditions>
 <action type="Rewrite" url="/" />
 </rule>
 </rules>
 </rewrite>
 </system.webServer>
</configuration>
```
9. Edit the **src\App.js** file to the following:
**src\App.js**
```js
import React, { Component } from 'react';
import './App.css';
import MenuContainer from './MenuContainer';


class App extends Component {
  render() {
    return (
      <div className="App">
        <MenuContainer />
      </div>
    );
  }
}

export default App;
```

10. Run these commands to build the project and create a ZIP file of the build folder:
```
npm run build
cd build
Compress-Archive -Path * -DestinationPath lunchmenu-build.zip
```
11. Open a web browser and go to https://YOUR_APP_NAME.scm.azurewebsites.net/ZipDeploy.
> **Note**: This is not the same site as your Azure Web App but requires the same credentials.

12. Drag the **build\lunchmenu-build.zip** file onto the file list in the browser and it will unzip and deploy the files to your web app.

![Azure deploy 2](../media/azure-deploy-3.png)

13. Browse to https://YOUR_APP_NAME.azurewebsites.net/ and confirm your React app is working.

![React lunchmenu 1](../media/react-lunchmenu-1.png)


## Convert React app to Team tab ##


1. In the terminal window in Visual Studio Code run these commands to install the packages for the Teams JS API and React SPA client-side routing:
```
npm install @microsoft/teams-js
npm install react-router-dom
```
> **Note:** If you aren't using npm, you can simple add a JS include to your HTML file:
```html
<script src="https://statics.teams.microsoft.com/sdk/v1.2/js/MicrosoftTeams.min.js" ></script>
```

2. Edit **src\App.js** to handle client routing for two pages.
```js
import React, { Component } from 'react';
import { NavLink, Switch, Route } from 'react-router-dom';

import './App.css';
import ConfigurePage from './ConfigurePage';
import MenuContainer from './MenuContainer';

const App = () => (
  <div className='app'>
   
     <Main />
  </div>
);

const Navigation = () => (
  <nav>
     <ul>
      <li><NavLink to='/'>Home</NavLink></li>
      <li><NavLink to='/config'>Config</NavLink></li>
    </ul> 
  </nav>
);

const Home = () => (
  <MenuContainer />
);


const Config = () => (
   <ConfigurePage />
);

const Main = () => (
  <Switch>
    <Route exact path='/' component={Home}></Route>
    <Route exact path='/config' component={Config}></Route>
  </Switch>
);

export default App;
```

3. Edit the **src\MenuContainer.js** file to get the user's UID upon loading.
```js
import React, { Component } from 'react';

import image_1 from './assets/barbecue-chicken.jpg';
import image_2 from './assets/tortellini.jpg';
import image_3 from './assets/eggplant.jpg';
import data_1 from "./assets/chicken_data.json";
import data_2 from "./assets/tortellini_data.json";
import data_3 from "./assets/eggplant_data.json";

import MenuItem from './MenuItem';   

import * as microsoftTeams from '@microsoft/teams-js';  //NEW: the Teams API include

class MenuContainer extends Component {
    /*NEW: create the username property in the component state*/
    constructor(props) {
        super(props);
        this.state = {userName: ""};   
    }   
    /*NEW: retrieve the username via async call to Teams API.  Lambda functions are needed to maintain the proper 'this' context.*/
    componentDidMount() {       
        microsoftTeams.initialize();    
        microsoftTeams.getContext(   
            (context) =>  { this.setState({ userName: context.upn}) } 
        );      
      }
    /*NEW: add the username to the title*/
    render(props) {
          return (
            <div className="MenuContainer">
                <header className="App-header">
                <h1 className="App-title">Hi {this.state.userName}, here is today's menu</h1>  
            </header>
                
            <MenuItem data_url={data_1} image_url={image_1}/>
            <MenuItem data_url={data_2} image_url={image_2}/>
            <MenuItem data_url={data_3} image_url={image_3}/>
            </div>
          );
 }
}
export default MenuContainer;
```
4. Edit the **src\index.js** file for client side routing:
```js
import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import registerServiceWorker from './registerServiceWorker';

import { BrowserRouter } from 'react-router-dom';  //NEW: the SPA router include

/* The SPA method for rendering the App component */
ReactDOM.render((
    <BrowserRouter>
      <App />
    </BrowserRouter>
    ), document.getElementById('root'));
registerServiceWorker();
```
5. Create a new file **src\ConfigurePage.js** in Visual Studio Code with the following:
```js
import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';            

class ConfigurePage extends Component {
    componentDidMount() {
        microsoftTeams.initialize();
    
        microsoftTeams.settings.setSettings({
            contentUrl: "https://YOUR_APP_NAME.azurewebsites.net/", // Mandatory parameter - the URL to show in the tab including query parameters if needed
            entityId: "MY_UNIQUE_IDENTIFIER", // Mandatory parameter - a unique ID or text for state tracking in the tab
            suggestedDisplayName: "Lunch Menu", //Optional parameter - the suggested text to display on the tab
			removeUrl: "https://YOUR_APP_NAME.azurewebsites.net/", // Optional parameter - the URL to show after a tab is removed
            websiteUrl: "https://YOUR_APP_NAME.azurewebsites.net/" //Optional parameter - the URL linked in the button “**Go to website**”
      });

       /* Register the function to call when the user clicks on the Save button when configuring a new tab */ 
      microsoftTeams.settings.registerOnSaveHandler(function(saveEvent){           
            saveEvent.notifySuccess();  // this call confirms a correct tab configuration and allows the process to continue.
        }); 
        
    microsoftTeams.settings.setValidityState(true);  // Activate the Save button when configuring a tab according to form validation. This form has no input, so it can be activated immediately.   
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
```

6. Run these commands to build the project and create a ZIP file of the build folder:
```
npm run build
cd build
Compress-Archive -Path * -DestinationPath lunchmenu-build.zip
```
7. Open a web browser and go to https://YOUR_APP_NAME.scm.azurewebsites.net/ZipDeploy.

8. Drag the update **build\lunchmenu-build.zip** file onto the file list in the browser and it will unzip and deploy the updated files to your web app.

9. Browse to https://YOUR_APP_NAME.azurewebsites.net/config and confirm your React app is working with the new configuration page.

![React lunchmenu config](../media/react-lunchmenu-config.png)


 
## Create an app manifest using Teams App Studio ##
1. Browse to https://teams.microsoft.com
2. Click on the **Store** icon and choose the **App Studio" app and click **Install**

![Teams app studio 1](../media/teams-appstudio-1.png)

3. Click on the **Open** button next to the App option to go to the App Studio

![Teams app studio 2](../media/teams-appstudio-2.png)

4. Change to the **Manifest editor** panel, click **Login** and accept the permissions requests.

![Teams app studio 3](../media/teams-appstudio-3.png)

5. Click on **+ Create a new app** and then on the new app card added in the app list.

![Teams app studio 4](../media/teams-appstudio-4.png)

6. Complete the manifest form with your app details.  Here are the two icons for the manifest: ![Lunch menu app icon 96x96](../media/lunchmenu_icon_96_96.png) and ![Lunch menu app icon 20x20](../media/lunchmenu_icon_20_20.png)

![Teams app studio 5](../media/teams-appstudio-5.png)
![Teams app studio 6](../media/teams-appstudio-6.png)
![Teams app studio 7](../media/teams-appstudio-7.png)
![Teams app studio 8](../media/teams-appstudio-8.png)

7. Click on the **Valid domains** selector from the left panel and add these Valid domains:
- secure.aadcdn.microsoftonline-p.com
- *.microsoftonline.com
- *.sharepointonline.com
- *.teams.microsoft.com

![Teams app studio 9](../media/teams-appstudio-9.png)


8. Click on the **Tabs** selector from the left panel under **Capabilities** and add a **Team tab** with this config URL: https://YOUR_APP_NAME.azurewebsites.net/config

![Teams app studio 10](../media/teams-appstudio-10.png)

9. Click on the **Test and Distribute** selector from the left panel under **Distribute** and then the **Export** button to download and save the **LunchMenu.zip** manifest file to your desktop.

![Teams app studio 10](../media/teams-appstudio-10.png)

## Sideload an app to a Team
> **Note**: For these steps, the Office 365 tenant must be configured to allow Teams app sideloading (uploading an app file instead of choosing from the Teams app store) 

1. Choose the Teams section and click on the ellipsis next to the **Lunch** team.

![Teams tab 1](../media/teams-tab-1.png)

2. Choose the **Apps* section and click on the **Upload a custom app** link in the bottom right corner.

![Teams tab 2](../media/teams-tab-2.png)

3. Locate and select the **Lunchmenu.zip** file previously downloaded and when uploaded the app will appear in the Team app list.

![Teams tab 3](../media/teams-tab-3.png)


## Add the custom app to a Team channel

1. Choose the **General** channel in the **Lunch** team and click on the **+** to add the new Lunch Menu app to the channel.

![Teams tab 4](../media/teams-tab-4.png)

![Teams tab 5](../media/teams-tab-5.png)

2. The configuration screen appears and renders the page returned from the configuration URL.
> **Note**: If the `microsoftTeams.settings.setValidityState(true);` not called in the configuration page, the **Save** button is never activated and the user cannot proceed.

![Teams tab 6](../media/teams-tab-6.png)

3. The custom tab is added to the channel and rendered correctly

![Teams tab 7](../media/teams-tab-7.png)


## Add Team UI components ##
To maintain the Teams and Office 365 look and feel, you can include the Teams CSS or React components to your project.  

Further details are in the **Control Library** in the Teams App Studio.

Color styling information is available here: [Teams - Design guidelines - color](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/design/components/color)


# Additional Reading topics #
### Bots ###
- [Create a bot with the Bot Builder SDK for .NET](https://docs.microsoft.com/en-us/azure/bot-service/dotnet/bot-builder-dotnet-quickstart?view=azure-bot-service-3.0)
- [Create a bot with the Bot Builder SDK for Node.js](https://docs.microsoft.com/en-us/azure/bot-service/nodejs/bot-builder-nodejs-quickstart?view=azure-bot-service-3.0)
- [Add bots to Microsoft Teams apps](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/bots/bots-overview)
- [Lab: Create a Bot in Teams](https://github.com/OfficeDev/TrainingContent/blob/master/Teams/03%20Bots/Lab.md)
- [Lab: Advanced Teams Bot](https://github.com/OfficeDev/TrainingContent/blob/master/Teams/05%20Microsoft%20Teams%20apps%20-%20Advanced%20Techniques/Lab.md#exercise1)

### Connectors ### 
- [Actionable messages in Outlook, Office 365 Groups, and Microsoft Teams](https://docs.microsoft.com/en-us/outlook/actionable-messages/#accessing-office-365-connectors-from-microsoft-teams)
- [Office 365 Connectors for Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/connectors)
- [Lab: Office 365 Connectors for Teams](https://github.com/OfficeDev/TrainingContent/blob/master/Teams/02%20Connectors/Lab.md)

### Excel add-ins ### 
- [Build an Excel add-in using React](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react)
- [Work with Charts using the Excel JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-charts)
- [Create custom JavaScript functions in Excel](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
- [Custom visuals in Power BI](https://docs.microsoft.com/en-us/power-bi/power-bi-custom-visuals)
- [Build Power BI Visuals](https://github.com/Microsoft/PowerBI-visuals) 

### O365 canvas example ### 

Further resources
•	Tabs dev
•	Teams Dev
