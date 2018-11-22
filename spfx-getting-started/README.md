# Global Office 365 Developer Bootcamp (2018)
### Hands-on Workshop - Development of an events list

## Setup an Office 365 Developer Tenant and Dev environment
If you are a developer, you can register for an Office 365 Developer tenant. There you can test the WebPart.

More information: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant

Ensure you installed NodeJS and the required global packages.

More information: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment

## Build the project
- Ensure NodeJS is installed; SPFx 1.6 requires NodeJS 8.x
- Ensure you have installed the global packages `gulp`, `yo` and `@microsoft/generator-sharepoint`
- Open the project folder within a console (VS Code Terminal, PowerShell, etc.) and run `npm i`
- Build a sppkg-File by running the commands `gulp bundle --ship` and `gulp package-solution --ship`

## Deploy the project
- Go to your SharePoint App Store. The URL for the App Store can look like this `https://{tenant-url}/sites/apps`
- Upload the file `./sharepoint/solution/event-list-web-part.sppkg` to the `Apps for SharePoint` library
- Trust the solution by clicking `Deploy`

![trust solution](./assets/05_trustapp.png)

## Approve Graph access
- Go to the new SharePoint Admin Center
- Under API management select the pending approval for `Calendars.ReadWrite` and click the `Approve or reject` button on the top of the page

![trust solution](./assets/06_approve_graph.png)

- Approve the app

## Create an event list
- Add a new list in the site collection you want to deploy the WebPart and name it `Events`
- Add 2 date fields to the list. One called `StartDate` and one called `EndDate`
- Add some items to the list

## Add WebPart to a page
- Go to a SharePoint modern page, click on `Site contents` and add the app `event-list-web-part-client-side-solution`
- After the app was installed, edit a modern page and add the WebPart `Event list`
- Edit the WebPart and set `Events` as list name
- Have fun

![WebPart](./assets/07_webpart.png)
