// Copyright 2001, 2002, 2003, 2004, 2005 Macromedia, Inc. All rights reserved.
<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function CreateVBArray(elem1,elem2,elem3,elem4)

	elem1 = "" + elem1
	elem2 = "" + elem2
	elem3 = "" + elem3
	elem4 = "" + elem4

	if (Len(elem1) = 0) then
		elem1 = Empty
	end if

	if (Len(elem2) = 0) then
		elem2 = Empty
	end if

	if (Len(elem3) = 0) then
		elem3 = Empty
	end if

	if (Len(elem4) = 0) then
		elem4 = Empty
	end if

	if (elem4 = "PrimaryKeys") then
		CreateVBArray = Array(elem1,elem2,elem3)
	else
		CreateVBArray = Array(elem1,elem2,elem3,elem4)
	end if

end function

function CreateVBEmpty()
	CreateVBEmpty = Empty
end function

</SCRIPT>


<SCRIPT LANGUAGE=JavaScript RUNAT=Server>

//define the variant array for openSchema call
//create a javascript array instead of vb variant
//to fix SP2
function CreateJSArray(elem1,elem2,elem3,elem4)
{
	var filterCriteriaArray = new Array();
		
	var filterParam1 = "" + elem1
	var filterParam2 = "" + elem2
	var filterParam3 = "" + elem3
	var filterParam4 = "" + elem4

	if (!((filterParam1.length > 0) && (filterParam1 != "undefined")))
	{
		filterParam1 = CreateVBEmpty();
	}
	
	if (!((filterParam2.length > 0) && (filterParam2 != "undefined")))
	{
		filterParam2 = CreateVBEmpty();
	}
	
	if (!((filterParam3.length > 0) && (filterParam3 != "undefined")))
	{
		filterParam3 = CreateVBEmpty();
	}
	
	if (!((filterParam4.length > 0) && (filterParam4 != "undefined")))
	{
		filterParam4 = CreateVBEmpty();
	}	
	
	if (filterParam4 == "PrimaryKeys") 
	{
		filterCriteriaArray[0] = filterParam1;
		filterCriteriaArray[1] = filterParam2;
		filterCriteriaArray[2] = filterParam3;		
	}
	else
	{
		filterCriteriaArray[0] = filterParam1;
		filterCriteriaArray[1] = filterParam2;
		filterCriteriaArray[2] = filterParam3;			
		filterCriteriaArray[3] = filterParam4;			
	}			
	return filterCriteriaArray;
}

function CreateMMConnection(ConnectionString,UserName,Password,Timeout)
{
	var Object;
	Object = new MMConnection(ConnectionString,UserName,Password,Timeout);
	return Object;
}

function MMConnection(ConnectionString,UserName,Password,Timeout)
{
	MMConnReconnect(this);
	this.isOpen = false;
	this.ConnectionString = ConnectionString;
	this.UserName		  = String(UserName);
	this.Password		  = String(Password);
	this.Connection		  = Server.CreateObject("ADODB.Connection");
	this.Connection.ConnectionTimeout = Timeout;
}


function MMConnReconnect(Object)
{
	Object.GetODBCDSNs				= ConnGetODBCDSNs;
	Object.Open						= ConnOpen;
	Object.GetTables				= ConnGetTables;
	Object.GetViews					= ConnGetViews;
	Object.GetProcedures			= ConnGetProcedures;
	Object.GetColumnsOfTable		= ConnGetColumns;
	Object.GetPrimaryKeysOfTable	= ConnGetPrimaryKeys;
	Object.GetParametersOfProcedure = ConnGetParametersOfProcedure;
	Object.ExecuteSQL				= ConnExecuteSQL;
	Object.ExecuteSP				= ConnExecuteSP;
	Object.ReturnsResultSet			= ConnReturnsResultSet;
	Object.SupportsProcedure		= ConnSupportsProcedure;
	Object.GetProviderTypes			= ConnGetProviderTypes;
	Object.HandleExceptions			= ConnHandleExceptions;
	Object.TestOpen					= ConnIsOpen;
	Object.Close					= ConnClose;
}


function ConnOpen()
{
	var theConnectionString = new String(this.ConnectionString);

	//  ????????????? OBSOLETE: begin ????????????????????????????????
	if (this.UserName && this.UserName.length)
	{
		theConnectionString = theConnectionString + ";uid=" + this.UserName;
	}
	if (this.Password && this.Password.length)
	{
		theConnectionString = theConnectionString + ";pwd=" + this.Password;
	}
	//  ????????????? OBSOLETE: end ????????????????????????????????

	//  The given connection string may not be formatted for OLE DB.  It may, for example,
	//  be a SQL Server connection string.  In such cases we need to morph it into
	//  an OLE DB connection string so it can be digested by the ADODB.Connection that
	//  we're using.
	//
	//  For now, we are only dealing with morphing SQL Server connection strings.  In the
	//  future, this logic may have to be expanded to deal with Oracle, Informix, etc. as
	//  those vendors make their own ASP.Net drivers available for use (circumventing
	//  the current need to go through OLE DB to access those databases).

	var dbType = Request("DATABASETYPE");
	if (dbType != null)
	{
		var strDBtype = new String(dbType);
		if ((strDBtype.length > 0) && (strDBtype.toLowerCase() == "sqlserver"))
		{
			if (theConnectionString.charAt(0) == "\"")
			{
				theConnectionString = "\"Provider=SQLOLEDB;" + theConnectionString.substring(1);
			}
			else
			{
				theConnectionString = "Provider=SQLOLEDB;" + theConnectionString;
			}
		}
	}

	var aConn = ConnEval(theConnectionString);
	this.Connection.Open(aConn);
	this.isOpen = (this.Connection.State == adStateOpen);
}

function ConnIsOpen()
{
	var xmlOutput = "";

	if (this.isOpen)
	{
		xmlOutput = xmlOutput + "<TEST status=";
		xmlOutput = xmlOutput + this.isOpen;
		xmlOutput = xmlOutput + "></TEST>";
	}

	return xmlOutput;
}

function ConnClose()
{
	if (this.Connection && this.isOpen)
	{
		this.Connection.Close();
	}
}

function ConnGetTables(SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,"","TABLE"));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,"","TABLE");		
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaTables,JSVariant));
	}
	
	return null;
}

function ConnGetViews(SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,"","VIEW"));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,"","VIEW");		
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaTables,JSVariant));
	}

	return null;
}

function ConnGetProcedures(SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,"",""));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,"","");				
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaProcedures,JSVariant));
	}

	return null;
}

function ConnGetColumns(TableName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,TableName,""));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,TableName,"");				
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaColumns,JSVariant));
	}

	return null;
}

function ConnGetPrimaryKeys(TableName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,TableName,"PrimaryKeys"));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,TableName,"PrimaryKeys");				
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaPrimaryKeys,JSVariant));
	}

	return null;
}


function ConnGetParametersOfProcedure(ProcedureName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,ProcedureName,""));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,ProcedureName,"");				
		return this.Connection.OpenSchema(adSchemaProcedureParameters,JSVariant);
	}

	return null;
}

function ConnExecuteSQL(aStatement,MaxRows)
{
	if (this.Connection && this.isOpen)
	{
		var oRecordset = Server.CreateObject("ADODB.Recordset");
		if (oRecordset)
		{
			aStatement = "" + aStatement;
			oRecordset.MaxRecords = MaxRows;
			oRecordset.Open(aStatement,this.Connection);
			return MarshallRecordsetIntoHTML(oRecordset);
		}
	}

	return null;
}

function ConnGetProviderTypes()
{
	if (this.Connection && this.isOpen)
	{
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaProviderTypes));
	}

	return null;
}

function ConnExecuteSP(aProcStatement,TimeOut,Parameters)
{
	if (this.Connection && this.isOpen)
	{
		var oCommand = Server.CreateObject("ADODB.Command");

		aProcStatement = "" + aProcStatement;
		oCommand.CommandTimeout = TimeOut;
		oCommand.CommandText = aProcStatement;
		oCommand.CommandType = adCmdStoredProc;
		oCommand.ActiveConnection = this.Connection;

		Parameters = "" + Parameters;

		if (!Parameters.length)
		{
			if (oCommand)
			{
				return MarshallRecordsetIntoHTML(oCommand.Execute());
			}
		}
		else
		{
			//Substitute Parameters.
			var Params = Parameters;
			var ParamArray = new Array();

			if (Params && Params != "undefined")
			{
				var cSize = 0;
				for (;;)
				{
					var index = Params.indexOf(",");
					if (index == -1)
					{
						index = Params.length;
					}

					var name = Params.substring(0,index);

					Params = Params.substring(index+1,Params.length);
					index = Params.indexOf(",");
					if (index == -1)
					{
						index = Params.length;
					}

					var value = Params.substring(0,index);

					var Pair = new Object();

					Pair.name = name;
					Pair.value = value;

					ParamArray[cSize] = Pair;
					cSize++;

					if (index >= Params.length)
					{
						break;
					}

					Params = Params.substring(index+1,Params.length);
				}


				if (oCommand.Parameters.Count == -1)
				{
					//Create Parameters
					var oRecordset = ConnGetParametersOfProcedure(aProcStatement);
					if (oRecordset)
					{
						var pCount=0;
						while (!oRecordset.EOF)
						{
							var pName    = oRecordset.Fields.Item("PARAMETER_NAME").Value;
							var pOrdinal = oRecordset.Fields.Item("ORDINAL_POSITION").Value;
							var pType	 = oRecordset.Fields.Item("PARAMETER_TYPE").Value;
							var pDataType = oRecordset.Fields.Item("DATA_TYPE").Value;
							switch (pDataType)
							{
								case adBinary:
								case adBSTR:
								case adChar:
								case adLongVarBinary:
								case adLongVarChar:
								case adLongVarWChar:
								case adLongVarChar:
								case adVarBinary:
								case adVarChar:
								case adVarWChar:
								{
									var pSize = oRecordset.Fields.Item("CHARACTER_MAXIMUM_LENGTH").Value;
								}
								default:
								{
									var pSize = null;
								}
							}

							if ((pType == adParamInput) || (pType == adParamInputOutput))
							{
								var pValue = ParamArray[pName];
								//if we could not find parameter by name ..try to find 
								//parameter by index.
								if (!pValue)
								{
									//try the case when the parameter is set by index.
									pStrCount = "" + pCount;
									pValue = ParamArray[pStrCount];
								}
								oCommand.CreateParameter(pName,pDataType,pType,pSize,pValue);
							}
							else
							{
								var pValue = null;
								oCommand.CreateParameter(pName,pDataType,pType,pSize,pValue);
							}
							oRecordset.MoveNext();
							pCount++;
						}
					}	
				 }
				 else
				 {
					for (var i =0 ; i < ParamArray.length ; i++)
					{
						Pair = ParamArray[i];

						if (Pair.value)
						{
							var pIndex = "" + parseInt(Pair.name);

							if (pIndex == Pair.name)
							{
								var aParameter = oCommand.Parameters(parseInt(Pair.name));
							}
							else
							{
								var aParameter = oCommand.Parameters(Pair.name);
							}

							if (aParameter)
							{
								if ((aParameter.Direction == adParamInput) || (aParameter.Direction == adParamInputOutput))
								{
									aParameter.Value = Pair.value;
								}
							}
						}
					}
					return MarshallRecordsetIntoHTML(oCommand.Execute());
				 }
			}
		}
	}

	return null;
}

function ConnReturnsResultSet(ProcedureName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		//var VBVariant =  new VBArray(CreateVBArray(CatalogName,SchemaName,ProcedureName,""));
		var JSVariant = CreateJSArray(CatalogName,SchemaName,ProcedureName,"");						
		var oRecordset = this.Connection.OpenSchema(adSchemaProcedureColumns,JSVariant);

		var status = "true";
		if (oRecordset.EOF) 
		{
			status = "false";
		}

		var xmlOutput = "";
		xmlOutput = xmlOutput + "<RETURNSRESULTSET status=";
		xmlOutput = xmlOutput + status;
		xmlOutput = xmlOutput + "></RETURNSRESULTSET>";
		return xmlOutput;
	}
}

function ConnSupportsProcedure()
{	
	if (this.Connection && this.isOpen)
	{
		var aProvider = "" + this.Connection.Provider;

		var status = "true";

		if (aProvider.indexOf("Microsoft.Jet") != -1)
		{
			status = "false";
		}

		if (aProvider.indexOf("MSDASQL")!=-1)
		{
			var ProviderTypes = this.Connection.OpenSchema(adSchemaProviderTypes);

			if (ProviderTypes.Fields.Count > 0)
			{
				//Access
				aProviderType = ProviderTypes.Fields(0).Value;
				aProviderType = aProviderType.toLowerCase();

				if (aProviderType == "guid")
				{
					status = "false";
				}//Paradox/DBaseIII.
				else if (aProviderType == "short")
				{
					status = "false";
				}
				else if (aProviderType == "image")
				{
					status = "false";
				}
				else if (aProviderType == "logical")
				{
					status = "false";
				} //For FoxPro
				else if (aProviderType == "l")
				{
					status = "false";
				} //For MySQL....
				else if (aProviderType == "tinyint")
				{
					status = "false";
				}
			}
		}

		var xmlOutput = "";
		xmlOutput = xmlOutput + "<SUPPORTSPROCEDURE status=";
		xmlOutput = xmlOutput + status;
		xmlOutput = xmlOutput + "></SUPPORTSPROCEDURE>";
		return xmlOutput;
	}
}

function ConnHandleExceptions()
{
	var xmlOutput = "";

	xmlOutput = xmlOutput + "<ERRORS>";
	if (this.Connection)
	{
		var Errors = this.Connection.Errors;

		for (var i =0 ; i < Errors.Count ; i++)
		{ 
			xmlOutput = xmlOutput + "<ERROR";

			xmlOutput = xmlOutput + " Identification=\""
			xmlOutput = xmlOutput + Errors(i).Number;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " Source=\""
			xmlOutput = xmlOutput + Errors(i).Source;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " HelpFile=\""
			xmlOutput = xmlOutput + Errors(i).HelpFile;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " HelpContext=\""
			xmlOutput = xmlOutput + Errors(i).HelpContext;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + "><DESCRIPTION>";
			xmlOutput = xmlOutput + HTMLEncode(Errors(i).Description);
			xmlOutput = xmlOutput + "</DESCRIPTION></ERROR>";
		}
	}
	xmlOutput = xmlOutput + "</ERRORS>";

	return xmlOutput;
}

function MarshallRecordsetIntoHTML(aResultSet)
{
	var xmlOutput = "";
	if (aResultSet)
	{
		xmlOutput = xmlOutput + "<RESULTSET>";
		xmlOutput = xmlOutput + "<FIELDS>";

		for(var i=0 ;i < aResultSet.Fields.Count ; i++)
		{
			xmlOutput = xmlOutput + "<FIELD";

			xmlOutput = xmlOutput + " type=\"";
			xmlOutput = xmlOutput + aResultSet.Fields(i).Type;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " definedSize=\"";
			xmlOutput = xmlOutput + aResultSet.Fields(i).DefinedSize;
			xmlOutput = xmlOutput + "\"";


			xmlOutput = xmlOutput + " actualsize=\"";

			if (!aResultSet.EOF)
			{
				xmlOutput = xmlOutput + aResultSet.Fields(i).ActualSize;
			}
			else
			{
				xmlOutput = xmlOutput + "-1";
			}

			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " precision=\"";
			xmlOutput = xmlOutput + aResultSet.Fields(i).Precision;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " scale=\"";
			xmlOutput = xmlOutput + aResultSet.Fields(i).NumericScale;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + ">";

			xmlOutput = xmlOutput + "<NAME>";
			xmlOutput = xmlOutput + HTMLEncode(aResultSet.Fields(i).Name);
			xmlOutput = xmlOutput + "</NAME>";
	
			xmlOutput = xmlOutput + "</FIELD>";

		}

		xmlOutput = xmlOutput + "</FIELDS>";
		xmlOutput = xmlOutput + "<ROWS>";

		while (!aResultSet.EOF)
		{
			xmlOutput = xmlOutput + "<ROW>";
			for(var i=0 ;i < aResultSet.Fields.Count ; i++)
			{
				xmlOutput = xmlOutput + "<VALUE>";
				var aValue = aResultSet.Fields(i).Value;
				if (aValue && aValue.length)
				{
					xmlOutput = xmlOutput + HTMLEncode(aValue);
				}
				else
				{
					xmlOutput = xmlOutput + aResultSet.Fields(i).Value;
				}
				xmlOutput = xmlOutput + "</VALUE>";
			}
			xmlOutput = xmlOutput + "</ROW>";
			aResultSet.MoveNext()
		}

		xmlOutput = xmlOutput + "</ROWS>";
		xmlOutput = xmlOutput + "</RESULTSET>";
		aResultSet.Close();
	}
	return xmlOutput;
}

function ConnGetODBCDSNs()
{
   var fso = new ActiveXObject("Scripting.FileSystemObject");
   var dsnList=new Array();
   var OdbcIniFile = null;
   var odbcFileName = "";
   var e = new Enumerator(fso.Drives);
   var xmlOutput="";
   for (; !e.atEnd(); e.moveNext())
   {
	  var x = e.item();

	  //Skip Drive that not ready...
	  if (!fso.DriveExists(x) || !x.IsReady || (x.DriveType==1))
	  {
		continue;
	  }

	  var driverLetter = x.DriveLetter;
	  var WinFolderName1 = driverLetter + ":\\" + "Winnt";
	  var WinFolderName2 = driverLetter + ":\\" + "Windows";
	  if (fso.FolderExists(WinFolderName1))
	  {
         //Get the ODBC FileName.
		odbcFileName = WinFolderName1 + "\\" + "ODBC.INI";
		break;
	  }
	  else if (fso.FolderExists(WinFolderName2))
	  {
         //Get the ODBC FileName.
		odbcFileName = WinFolderName2 + "\\" + "ODBC.INI";
		break;
	  }
   }

   if (odbcFileName.length > 0)
   {
	  if (fso.FileExists(odbcFileName))
	  {
	     //  Don't use the FSO's OpenTextFile method because it hangs Windows XP.
		 //  Work around that Windows bug by using the equivalently functional Stream
		 //  object.

		 OdbcIniFile = Server.CreateObject("ADODB.Stream");
		 OdbcIniFile.Type = 2;

		 //  Initially, we will try to use the charset associated with the codepage
		 //  for this ASP session.  This should be the default charset for this
		 //  web site or for this server.  This helps in cases where the OS is in,
		 //  say, Japanese.  However, we have found (empiracally) that on some
		 //  computers this causes the stream to be unreadable.  So, we will
		 //  (below) test the stream.  If it looks bogus, we will default to using
		 //  "ascii" as the charset.

		 OdbcIniFile.Charset = getCharsetStringFromCodepageNumber(currentSessionCodePage);
		 OdbcIniFile.Open();
		 OdbcIniFile.LoadFromFile(odbcFileName);

		 if (OdbcIniFile && (!OdbcIniFile.EOS))
		 {
			 //  Test to see if we can, in fact, read the stream.

			 var aLine = OdbcIniFile.ReadText(-2);

			 //  Reset the stream's position so if we can read it we haven't lost the test line
			 //  that we just read.

			 OdbcIniFile.Position = 0;

			 if (aLine == "")
			 {
				 //  It appears that the stream is bogus.  Try reading it as ASCII.

				 OdbcIniFile.Close();
				 OdbcIniFile.Type = 2;
				 OdbcIniFile.Charset = "ascii";
				 OdbcIniFile.Open();
				 OdbcIniFile.LoadFromFile(odbcFileName);
			 }
		 }
      }
   }

   if (OdbcIniFile)
   {
	 var i =0;
	 var odbcSection = -1;
	 while (!OdbcIniFile.EOS)
	 {
		 var aLine = OdbcIniFile.ReadText(-2);
		 var odbcSection = aLine.indexOf("[ODBC");
		 if (odbcSection != -1)
		 {
			break;
		 }
	 }
	 if (odbcSection != -1)
	 {
		 while (!OdbcIniFile.EOS)
		 {
			 var aLine = OdbcIniFile.ReadText(-2);
			 if (aLine.charAt(0) != "[")
			 {
				 var anIndex = aLine.indexOf("=");
				 if (anIndex != -1)
				 {
					var dsnName = aLine.substring(0,anIndex);
					dsnList[dsnList.length]= dsnName;
				 }
			}
			else
			{
				break;
			}
		 }
	  }
	 OdbcIniFile.Close();
   }

   xmlOutput = "<RESULTSET>";
   if (dsnList.length)
   {
		xmlOutput = xmlOutput + "<FIELDS>";
		xmlOutput = xmlOutput + "<FIELD>";
		xmlOutput = xmlOutput + "<NAME>NAME</NAME>";
		xmlOutput = xmlOutput + "</FIELD>";
		xmlOutput = xmlOutput + "</FIELDS>";

		xmlOutput = xmlOutput + "<ROWS>";
		for (var i =0 ; i < dsnList.length; i++)
		{
			xmlOutput = xmlOutput + "<ROW>";
			xmlOutput = xmlOutput + "<VALUE>";
			xmlOutput = xmlOutput + HTMLEncode(dsnList[i]);
			xmlOutput = xmlOutput + "</VALUE>";
			xmlOutput = xmlOutput + "</ROW>";
		}
		xmlOutput = xmlOutput + "</ROWS>";
   }
   xmlOutput = xmlOutput + "</RESULTSET>";
   return xmlOutput;
}

function ConnEval(ConnString)
{
	ConnString = "" + ConnString;
	if (ConnString.length)
	{
		var delimiter = (ConnString.indexOf("+") != -1) ? "+" : "&";
		var aConnString = "";

		for (;;)
		{
			var index = ConnString.indexOf(delimiter);

			if (index == -1)
			{
				index = ConnString.length;
			}

			var aStringlet	= ConnString.substring(0,index);

			try
			{
				aConnString = aConnString + eval(aStringlet);
			}
			catch (e)
			{
				aConnString = ConnString;
				return aConnString;
			}

			if (index >= ConnString.length)
			{
				break;
			}

			ConnString = ConnString.substring(index+1,ConnString.length);
		}

		return aConnString;
	}

	return ConnString;
}

function HTMLEncode(TheString)
{
	if ( Session.CodePage == 65001 )
	{
		return TheString;
	}
	else
	{
		return Server.HTMLEncode(TheString);
	}
}

function getCharsetStringFromCodepageNumber(nCodePage)
{
	var strCharSet = "ascii";

	switch (nCodePage)
	{
		case 20106: strCharSet = "DIN_66003"; break;
		case 20108: strCharSet = "NS_4551-1"; break;
		case 20107: strCharSet = "SEN_850200_B"; break;
		case 50932: strCharSet = "_autodetect"; break;
		case 50949: strCharSet = "_autodetect_kr"; break;
		case 950: strCharSet = "big5"; break;
		case 50221: strCharSet = "csISO2022JP"; break;
		case 51949: strCharSet = "euc-kr"; break;
		case 936: strCharSet = "gb2312"; break;
		case 52936: strCharSet = "hz-gb-2312"; break;
		case 852: strCharSet = "ibm852"; break;
		case 866: strCharSet = "ibm866"; break;
		case 20105: strCharSet = "irv"; break;
		case 50220: strCharSet = "iso-2022-jp"; break;
		case 50222: strCharSet = "iso-2022-jp"; break;
		case 50225: strCharSet = "iso-2022-kr"; break;
		case 1252: strCharSet = "iso-8859-1"; break;
		case 28591: strCharSet = "iso-8859-1"; break;
		case 28592: strCharSet = "iso-8859-2"; break;
		case 28593: strCharSet = "iso-8859-3"; break;
		case 28594: strCharSet = "iso-8859-4"; break;
		case 28595: strCharSet = "iso-8859-5"; break;
		case 28596: strCharSet = "iso-8859-6"; break;
		case 28597: strCharSet = "iso-8859-7"; break;
		case 28598: strCharSet = "iso-8859-8"; break;
		case 20866: strCharSet = "koi8-r"; break;
		case 949: strCharSet = "ks_c_5601"; break;
		case 932: strCharSet = "shift-jis"; break;
		case 1200: strCharSet = "unicode"; break;
		case 1201: strCharSet = "unicodeFEFF"; break;
		case 65000: strCharSet = "utf-7"; break;
		case 65001: strCharSet = "utf-8"; break;
		case 1250: strCharSet = "windows-1250"; break;
		case 1251: strCharSet = "windows-1251"; break;
		case 1252: strCharSet = "windows-1252"; break;
		case 1253: strCharSet = "windows-1253"; break;
		case 1254: strCharSet = "windows-1254"; break;
		case 1255: strCharSet = "windows-1255"; break;
		case 1256: strCharSet = "windows-1256"; break;
		case 1257: strCharSet = "windows-1257"; break;
		case 1258: strCharSet = "windows-1258"; break;
		case 874: strCharSet = "windows-874"; break;
		case 51932: strCharSet = "x-euc"; break;
		case 50000: strCharSet = "x-user-defined"; break;
	}

	return strCharSet;
}

</SCRIPT>

<SCRIPT LANGUAGE=JavaScript RUNAT=Server>

//---- ObjectStateEnum Values ----
var adStateClosed = 0x00000000;
var adStateOpen = 0x00000001;
var adStateConnecting = 0x00000002;
var adStateExecuting = 0x00000004;
var adStateFetching = 0x00000008;

//---- DataTypeEnum Values ----
var adEmpty = 0;
var adTinyInt = 16;
var adSmallInt = 2;
var adInteger = 3;
var adBigInt = 20;
var adUnsignedTinyInt = 17;
var adUnsignedSmallInt = 18;
var adUnsignedInt = 19;
var adUnsignedBigInt = 21;
var adSingle = 4;
var adDouble = 5;
var adCurrency = 6;
var adDecimal = 14;
var adNumeric = 131;
var adBoolean = 11;
var adError = 10;
var adUserDefined = 132;
var adVariant = 12;
var adIDispatch = 9;
var adIUnknown = 13;
var adGUID = 72;
var adDate = 7;
var adDBDate = 133;
var adDBTime = 134;
var adDBTimeStamp = 135;
var adBSTR = 8;
var adChar = 129;
var adVarChar = 200;
var adLongVarChar = 201;
var adWChar = 130;
var adVarWChar = 202;
var adLongVarWChar = 203;
var adBinary = 128;
var adVarBinary = 204;
var adLongVarBinary = 205;
var adChapter = 136;
var adFileTime = 64;
var adDBFileTime = 137;
var adPropVariant = 138;
var adVarNumeric = 139;

//---- PositionEnum Values ----
var adPosUnknown = -1;
var adPosBOF = -2;
var adPosEOF = -3;

//---- ParameterDirectionEnum Values ----
var adParamUnknown = 0x0000;
var adParamInput = 0x0001;
var adParamOutput = 0x0002;
var adParamInputOutput = 0x0003;
var adParamReturnValue = 0x0004;

//---- CommandTypeEnum Values ----
var adCmdUnknown = 0x0008;
var adCmdText = 0x0001;
var adCmdTable = 0x0002;
var adCmdStoredProc = 0x0004;
var adCmdFile = 0x0100;
var adCmdTableDirect = 0x0200;


//---- SchemaEnum Values ----
var adSchemaProviderSpecific = -1;
var adSchemaAsserts = 0;
var adSchemaCatalogs = 1;
var adSchemaCharacterSets = 2;
var adSchemaCollations = 3;
var adSchemaColumns = 4;
var adSchemaCheckConstraints = 5;
var adSchemaConstraintColumnUsage = 6;
var adSchemaConstraintTableUsage = 7;
var adSchemaKeyColumnUsage = 8;
var adSchemaReferentialConstraints = 9;
var adSchemaTableConstraints = 10;
var adSchemaColumnsDomainUsage = 11;
var adSchemaIndexes = 12;
var adSchemaColumnPrivileges = 13;
var adSchemaTablePrivileges = 14;
var adSchemaUsagePrivileges = 15;
var adSchemaProcedures = 16;
var adSchemaSchemata = 17;
var adSchemaSQLLanguages = 18;
var adSchemaStatistics = 19;
var adSchemaTables = 20;
var adSchemaTranslations = 21;
var adSchemaProviderTypes = 22;
var adSchemaViews = 23;
var adSchemaViewColumnUsage = 24;
var adSchemaViewTableUsage = 25;
var adSchemaProcedureParameters = 26;
var adSchemaForeignKeys = 27;
var adSchemaPrimaryKeys = 28;
var adSchemaProcedureColumns = 29;
var adSchemaDBInfoKeywords = 30;
var adSchemaDBInfoLiterals = 31;
var adSchemaCubes = 32;
var adSchemaDimensions = 33;
var adSchemaHierarchies = 34;
var adSchemaLevels = 35;
var adSchemaMeasures = 36;
var adSchemaProperties = 37;
var adSchemaMembers = 38;
</SCRIPT>


