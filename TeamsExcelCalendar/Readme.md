# Microsoft Teams Excel Calendar Tab sample

The following sample is a Microsoft Teams tab application that will create an Excel Spreadsheet based on one or more Exchange Online Mailbox Calendars by calling the Microsoft Graph API  to get the Calendar data and build the spreadsheet

Screen Shots of the Tab application in Action 

![enter image description here](https://1.bp.blogspot.com/-wn9H6H5cBLk/Xdzx_8kjvGI/AAAAAAAAChU/RVyunVxRm8QWWJ9x9VWwq8ofFigv0mmKwCLcBGAsYHQ/s1600/agCal.JPG)

# Exchange Online Mailbox Permission requirements #

As this app runs as the currently logged on user that user must have at least reviewer access to the calendars that are going to be aggregated.

# Configuration of the Calendars to Aggregate #
This Tab app will read a JSON configuration file from the Team's Channel Document library. The file needs to be called ExcelCalendarConfig.json. With the following format

    {
    "Calendars": [
        {
            "CalendarEmailAddress": "SharedMailbox@datarumble.com",
            "CalendarName": "Australia holidays",
            "CalendarDisplayName": "Australia"
        },
        {
            "CalendarEmailAddress": "SharedMailbox@datarumble.com",
            "CalendarName": "United States holidays",
            "CalendarDisplayName": "United States"
        },
        {
            "CalendarEmailAddress": "SharedMailbox@datarumble.com",
            "CalendarName": "United Kingdom holidays",
            "CalendarDisplayName": "UK"
        }
    ]
}
The difference between the CalendarName and CalendarDisplayName is  the CalendarDisplayName is how you want this calendar reflected in the spreadsheet vs that CalendarName which needs to be the exact name of the calendar to query. To install the config file just use the Teams client to drag and drop it in the Root of the Files tab for the channel you going to install the app onto.

# **Installation** #

**Prerequisites for using Teams Tab Applications**

To use a Teams Tab application application side loading must be enabled in the Office365 portal see the following page for how to modify the Teams Org setting [https://docs.microsoft.com/en-us/microsoftteams/admin-settings](https://docs.microsoft.com/en-us/microsoftteams/admin-settings). 
> "Sideloading is how you add an app to Teams by uploading a zip file directly to a team. Side-loading lets you test an app as it's being developed. It also lets you build an app for internal use only and share it with your team without submitting it to the Teams app catalog in the Office Store. "

![](https://gscales.github.io/TeamsGroupCalendar/docs/Sideloading.JPG)

**Note**: Make sure you use the https://admin.microsoft.com/AdminPortal/Home#/Settings/ServicesAndAddIns and not the Teams Admin portal as you won't be able to finding this setting in the later.

# Testing this GitHub Instance #

The application files for a Teams Tab application needs to be hosted on a web server, for testing only you can use this hosted version on gitHub. To use this you would need to grant the following applicationId consent in your tenant using the following URL

[https://login.microsoftonline.com/common/adminconsent?client_id=749e9a57-fbc5-4364-a01b-d93a68a640ce](https://login.microsoftonline.com/common/adminconsent?client_id=749e9a57-fbc5-4364-a01b-d93a68a640ce)

You then need to download the Manifest Zip file from [https://github.com/gscales/gscales.github.io/raw/master/TeamsExcelCalendar/TabPackage/app.zip
](https://github.com/gscales/gscales.github.io/raw/master/TeamsExcelCalendar/TabPackage/app.zip)
then follow the Custom App installation process described below


# **Custom App Installation Process** #

Official documentation for installing Custom Apps can be found 
[https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload)

Walk through

Choose the Microsoft Store Icon in the Team client (don't worry you not going to purchase anything)
![](https://gscales.github.io/TeamsGroupCalendar/docs/walkthrough1.JPG)

Then select "Upload Custom Application" & "For Me and My Teams"

Select the Team you want to install the app into 

Then select the Channel to install the Tab onto.

# Hosting the Application yourself #

Create an Azure AD Application Registration that has the following grants
![enter image description here](https://gscales.github.io/TeamsExcelCalendar/app/Docs/permsforce.JPG)


Modify the manifest of the Application registration to enable the Implicit authentication flow 

    "logoutUrl": null,
  	"oauth2AllowImplicitFlow": true,
    "oauth2AllowUrlPathMatching": false,

Change the tab application m** Manifest** (your version of [https://github.com/gscales/gscales.github.io/blob/master/TeamsExcelCalendar/TabPackage/manifest.json](https://github.com/gscales/gscales.github.io/blob/master/TeamsExcelCalendar/TabPackage/manifest.json))

You need to change the Id,PackageName and configurationURL setting in the manifest to your own unique ApplicationId and URL where the config.html page is hosted

      "$schema": "https://statics.teams.microsoft.com/sdk/v1.2/manifest/MicrosoftTeams.schema.json", 
  	  "manifestVersion": "1.5",
      "version": "1.0.0",
      "id": "749e9a57-fbc5-4364-a01b-d93a68a640ce",
      "packageName": "TeamsExcelCalendar.io.github.gscales",
    configurableTabs": [
    {
      "configurationUrl": "https://gscales.github.io/TeamsExcelCalendar/app/config.html",
      "canUpdateConfiguration": true,
      "scopes": [ "team" ]
    }
    ],

Modify you hosted version of the https://github.com/gscales/gscales.github.io/blob/master/TeamsExcelCalendar/app/Config/appconfig.js file. Change the clientId to the applicationId from your application registration and the hostRoot to the root of your webhost.

     const getConfig = () => {
  	 var config = {
        clientId : "749e9a57-fbc5-4364-a01b-d93a68a640ce",
        redirectUri : "/TeamsExcelCalendar/app/silent-end.html",
        authwindow :  "/TeamsExcelCalendar/app/auth.html",
	 hostRoot: "https://gscales.github.io",
   	 };
  	 return config;
	}

Create a Zip file of all the files in the https://github.com/gscales/gscales.github.io/blob/master/TeamsExcelCalendar/TabPackage directory and then use that in the Custom App Installation process described above.













