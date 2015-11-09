Ticket Mailer
====================

### Configuration

open \Emailer\bin\Release\Emailer.exe.config

This is an xml document that the application scans to build its configuration.  In the <applicationSettings> node there are nine settings that are editable.  If the document is modified in such a way that it is no longer parsable as xml, the application will not function.

Prepend: This text will be prepended to the email before the first ticket field.

HelpDeskEmail: This is the email the application will send to.

HelpDeskHost: The server the email will be sent to.  Typically stmp.{your site}.com or an ip address.

NoInternetPhone: If the application is unable to ping the google servers, it is assumed that it has no internet connection and will display this value.

HostUserName: The username of the user on your smtp server you want to use

HostPassword: The password of your stmp username

IssueSeverityOptions: Each value within the <string> tags in the <ArrayOfString...> node is an option for issue severity.  Currently the application only supports numbers for these values.

OnlineSupportLink: A link to display at the bottom of the application main screen

BrandImagePath: The path to the image that is displayed in the top right of the main application screen



### Modifying the Application

Using any version of Visual Studio 2015, open Emailer.sln in the root directory.  This will load the solution.

Right click on DeusEmailer.cs and click View Code.  This will open the c# that powers the application.  By default, Visual Studio will attempt to load the designer view rather than the code if you just double click the file.

Once your modifications are complete you can run the application from within Visual Studio by clicking the Start (with a green arrow next to it) button in your taskbar.

To build the project for release, go to Build -> Batch Build.  You will have options to build both a Debug version and a Release version.  For standard usage, you should only need the release build.  You can change where the application builds in the properties of the project (not the solution).  Right click the project (c# Emailer) in your solution explorer and select properties to access this menu.