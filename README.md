#asp-airbrake

Classic ASP integration for [Airbrake.io](http://airbrake.io) because some devs still have to support Classic ASP apps, so why shouldn't they get to play with the latest toys?

##Usage (already handling errors)

Include the inc file and call the Airbrake class with your API key and environment name

    <!-- #include file="airbrake.inc.asp" -->
    <% Set AirBrake = (New AirBrakeClass)("YOUR-KEY", "production") %>

If you want to continue processing errors in your original way, the class exposes the original error object, access it like this:

    Set objError = AirBrake.GetLastError()

##Usage (from scratch)
If you are not already handling your ASP errors, why not?

The simplest method is to create a new page (e.g. `500.asp`) and insert the following:

    <!-- #include file="airbrake.inc.asp" -->
    <% Set AirBrake = (New AirBrakeClass)("YOUR-KEY", "production") %>
    <% Response.Clear %>
    <html>
    <head>
    <title>There was an error</title>
    </head>
    <body>
    <h1>Server Error</h1>
    <p>There was an error on the site, however a log has been taken and the issue will be looked at as soon as possible</p>
    </body>
    </html>

Then add the following to your `web.config`

	<?xml version="1.0" encoding="utf-8"?>
	<configuration>
	    <system.webServer>
			<httpErrors errorMode="DetailedLocalOnly">
	            <error statusCode="500" subStatusCode="100" path="500.asp" responseMode="ExecuteURL" />
	        </httpErrors>
	    </system.webServer>
	</configuration>
	
Adjust paths accordingly.