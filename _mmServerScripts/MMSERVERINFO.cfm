<cfprocessingdirective pageencoding="utf-8">
<cfcontent type="text/xml; charset=utf-8">
<cfset setEncoding("URL", "utf-8")>
<cfset setEncoding("FORM", "utf-8")>
<ServerInfo>
	<!--- Return the ColdFusion version several different ways for backwards compatability because 
		  the ColdFusion server model uses "Cold Fusion" with a space and "ColdFusion" witout 
		  interchangedly. For the rest of the values use their full name so it's clear what the 
		  value is and where it came from. Extension Authors can access the values returned from
		  this file using the API dom.serverModel.getServerVersion(), for example
		  dom.serverModel.getServerVersion("Server.ColdFusion.ProductLevel") to see if this is
		  standard or enterprise. Dreamweaver only requests this file ones per application session
		  and keeps a local copy of the response in the users site cache folder --->
	<ServerVersion>
		<Name>ColdFusion</Name>
		<Value><cfoutput>#Server.ColdFusion.ProductVersion#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Cold Fusion</Name>
		<Value><cfoutput>#Server.ColdFusion.ProductVersion#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductVersion</Name>
		<Value><cfoutput>#Server.ColdFusion.ProductVersion#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductVersion.Major</Name>
		<Value><cfoutput>#ListGetAt(Server.ColdFusion.ProductVersion, 1)#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductVersion.Minor</Name>
		<Value><cfoutput>#ListGetAt(Server.ColdFusion.ProductVersion, 2)#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductVersion.Patch</Name>
		<Value><cfoutput>#ListGetAt(Server.ColdFusion.ProductVersion, 3)#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductVersion.Build</Name>
		<Value><cfoutput>#ListGetAt(Server.ColdFusion.ProductVersion, 4)#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductLevel</Name>
		<Value><cfoutput>#Server.ColdFusion.ProductLevel#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.ProductName</Name>
		<Value><cfoutput>#Server.ColdFusion.ProductName#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.SupportedLocales</Name>
		<Value><cfoutput>#Server.ColdFusion.SupportedLocales#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.AppServer</Name>
		<Value><cfoutput>#Server.ColdFusion.AppServer#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.ColdFusion.InstallKit</Name>
		<Value><cfoutput>#Server.ColdFusion.InstallKit#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.OS.Version</Name>
		<Value><cfoutput>#Server.OS.Version#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.OS.BuildNumber</Name>
		<Value><cfoutput>#Server.OS.BuildNumber#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.OS.Name</Name>
		<Value><cfoutput>#Server.OS.Name#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.OS.Arch</Name>
		<Value><cfoutput>#Server.OS.Arch#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>Server.OS.AdditionalInformation</Name>
		<Value><cfoutput>#Server.OS.AdditionalInformation#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.GATEWAY_INTERFACE</Name>
		<Value><cfoutput>#CGI.GATEWAY_INTERFACE#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.HTTP_ACCEPT</Name>
		<Value><cfoutput>#CGI.HTTP_ACCEPT#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.HTTP_ACCEPT_LANGUAGE</Name>
		<Value><cfoutput>#CGI.HTTP_ACCEPT_LANGUAGE#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.HTTP_HOST</Name>
		<Value><cfoutput>#CGI.HTTP_HOST#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.SERVER_NAME</Name>
		<Value><cfoutput>#CGI.SERVER_NAME#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.SERVER_PORT</Name>
		<Value><cfoutput>#CGI.SERVER_PORT#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.SERVER_PORT_SECURE</Name>
		<Value><cfoutput>#CGI.SERVER_PORT_SECURE#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.SERVER_PROTOCOL</Name>
		<Value><cfoutput>#CGI.SERVER_PROTOCOL#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.SERVER_SOFTWARE</Name>
		<Value><cfoutput>#CGI.SERVER_SOFTWARE#</cfoutput></Value>
	</ServerVersion>
	<ServerVersion>
		<Name>CGI.WEB_SERVER_API</Name>
		<Value><cfoutput>#CGI.WEB_SERVER_API#</cfoutput></Value>
	</ServerVersion>
</ServerInfo>
