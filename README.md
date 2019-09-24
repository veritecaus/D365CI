
# Dynamics 365 Continuous Integration

This project provides PowerShell cmdlets that can be used to automate the importing and exporting of Dynamics 365 solutions and reference data.

Migrating Dynamics 365 solutions and reference data from one environment to another is not a trivial task. The Microsoft Dynamics team supply enough tools to get the basic job done, however if you've ever tried to achieve full automation for complex configurations then you'll know how complex and time consuming it can be.

There are a number of products and community driven projects in the market to help try and solve this problem. However for varying reasons they all have their limitations. 

The high level goals of this software is to:
-   Automate Dynamics 365 solution deployment
-   Automate Dynamics 365 reference data deployment
-   Easy to use and fast to implement for new projects
-   Provide artifacts that can be tracked using source control
-   Ability to run as part of build manager (eg Azure DevOps Build Server)

# Getting Started
The following guide should be enough to get you importing and exporting solution and reference data using PowerShell.

## Download & compile

Acquire the source from this repository and compile. Source has been tested to compile with Visual Studio 2017. Once compiled, you should have the following assemblies:

```powershell
Veritec.Dynamics.CI.PowerShell.dll
 
DocumentFormat.OpenXml.dll
Microsoft.IdentityModel.Clients.ActiveDirectory.dll
Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll
Microsoft.Xrm.Sdk.Deployment.dll
Microsoft.Xrm.Sdk.dll
Microsoft.Xrm.Tooling.Connector.dll
System.Management.Automation.dll
Veritec.Dynamics.CI.Common.dll
```

##  Export & Import Dynamics 365 Solution
Use the below PowerShell to export your solution from Dynamics 365. Make sure you change the **ConnectionString** and **SolutionName** parameters. SolutionName is a semicolon delimited list of the solutions you would like to export.

```powershell
# Reference PowerShell module
import-module "C:\VSRC\Veritec.Dynamics.CI\Veritec.Dynamics.CI.PowerShell\bin\Release\Veritec.Dynamics.CI.PowerShell.dll"
 
#### 1. Office 365 Method for logging in
$currentUser = Get-Credential -UserName "youruser@yourorg.onmicrosoft.com" -Message "D365 Credentials:"
$userName = $currentUser.UserName
$password = $currentUser.GetNetworkCredential().Password

$connectString =  "AuthType=Office365;Url=https://yourorg.crm6.dynamics.com;RequireNewInstance=True;UserName=$userName;Password=$password"

#### 2. S2S method Method for logging in (Note: AppId requires an app registration in Azure AD.)
#$connectString = "AuthType=OAuth;Url=https://yourorg.crm6.dynamics.com;AppId=yourAppIDGuid;UserName=youruser@yourorg.onmicrosoft.com;RedirectUri=https://yourorg.crm6.dynamics.com;LoginPrompt=Always;TokenCacheStorePath=c:\temp\mytoken;RequireNewInstance=True;"
 
# Export Dynamics Solution from source
Export-DynamicsSolution `
          -ConnectionString $connectString `
          -SolutionName "Solution1;Solution2;" `
          -SolutionDir "..\LocalSolutionDirectory" `
```
Add this extra statement to import the solution to your target tenant. Make sure you change the **ConnectionString** to match your target environment

```powershell
# Import Dynamics Solution to target
Import-DynamicsSolution `
          -ConnectionString $connectString `
          -SolutionName "Solution1;Solution2" `
          -SolutionDir "..\LocalSolutionDirectory" `
```

##  Export & Import Dynamics 365 Data
The following powershell commandlet can be used to migrate data from one tenant to another, which includes some system data like business units, queues, SLAs, Word Templates etc.

Use the below PowerShell to export your reference data from Dynamics 365. Make sure you change the  **ConnectionString**  and **FetchXMLFile** parameters. The FetchXMLFile parameter defines the location of the XML file that contains FetchXML queries that you would like to use to extract your reference data.
```powershell
# Export Dynamics Data
Export-DynamicsData `
      -ConnectionString $connectString `
      -FetchXMLFile ".\FetchXMLQueries.xml" `
      -OutputDataPath ".\ReferenceData"    
```
Here is an example **FetchXMLFile** file containing queries that we recommend for every single Dynamics 365 solution migration as a minimum. Save this file to disk and reference it in the above PowerShell. It will export the data for "business units", "currencies" and "teams".

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
 
<FetchXMLQueries>
  <FetchXMLQuery name="business units" order="00010">
    <fetch>
      <entity name="businessunit" >
        <attribute name="businessunitid" />
        <attribute name="name" />
        <attribute name="address1_addressid" />
        <attribute name="parentbusinessunitid" />
        <attribute name="address2_addressid" />
        <filter>
          <condition attribute="parentbusinessunitidname" operator="not-null" />
        </filter>
      </entity>
    </fetch>
  </FetchXMLQuery>
 
  <FetchXMLQuery name="currencies" order="00012">
    <fetch>
      <entity name="transactioncurrency">
        <attribute name="transactioncurrencyid" />
        <attribute name="currencyname" />
        <attribute name="isocurrencycode" />
        <attribute name="currencysymbol" />
        <attribute name="exchangerate" />
        <attribute name="currencyprecision" />
        <attribute name="statecode" />
      </entity>
    </fetch>
  </FetchXMLQuery>
 
  <!--There may be some workflow that can set the owner of a record to a team - so migrate team before importing any D365 solution -->
  <FetchXMLQuery name="teams" order="00020">
    <fetch>
      <entity name="team" >
        <attribute name="teamid" />
        <attribute name="name" />
        <attribute name="businessunitid" />
        <attribute name="administratorid" />
        <attribute name="teamtype" />
        <order attribute="name" descending="false" />
        <filter>
          <condition attribute="isdefault" operator="eq" value="0" />
        </filter>
      </entity>
    </fetch>
  </FetchXMLQuery>
</FetchXMLQueries>
 ```

Once you're able to export your data above. Add this extra statement to import the reference data to a target tenant. Make sure you change the **ConnectionString** and **TransformFile** parameters to match your target environment.
```powershell
# Import Dynamics Data
Import-DynamicsData `
    -ConnectionString $connectString `
    -EncryptedPassword $encryptedPwd `
    -TransformFile "transforms.json"

Import-DynamicsData `
    -ConnectionString $connectString `
    -TransformFile "transforms.json" `
    -InputDataPath ".\ReferenceData"
```
The Transform file is used to modify your data to be inserted into your target environment. This is useful when:

1.  The GUID of an object is different in a target environment and you can't change it - eg the default root business unit GUID which is created/generated when the tenant is created
2.  You have a value in a target environment that you would like to be different on purpose - eg some of the ADX Studio Site Settings (adx_sitesetting) values.

Here is a minimum set of transforms that we recommend for every single Dynamics 365 solution migration. Save this file to disk and reference it in the above PowerShell. It will replace the GUID for the root business unit and any other business units that refer to it as a parent. Make sure you change "407C6265-9742-E811-A94F-000D3AD064BD" with the GUID of your root business unit in the source, and "9CDB1F7D-2F2A-E811-A853-000D3AD07676" with the GUID or the root business unit in your target. Also replace the organization GUID.

```json
[
  {
    TargetEntity: "businessunit",
    TargetAttribute : "parentbusinessunitid",
    TargetValue : "*",
    ReplacementValue: "407C6265-9742-E811-A94F-000D3AD064BD"
   },
   {
     TargetEntity: "businessunit",
     TargetAttribute: "businessunitid",
     TargetValue: "9CDB1F7D-2F2A-E811-A853-000D3AD07676",
     ReplacementValue : "407C6265-9742-E811-A94F-000D3AD064BD"
   }
]
```
##  Bring it All Together
Now that you know how to export and import your solution and reference data. Now it's time to explore what this would look like in a production scenario.

Example PowerShell to export Dynamics Solution and Reference Data. Notice how there is two sets of data being exported, one for pre and one for post solution. This is often necessary as some data needs to be made available before a solution is imported (eg currency). 

```powershell
# Reference PowerShell module
import-module "C:\VSRC\Veritec.Dynamics.CI\Veritec.Dynamics.CI.PowerShell\bin\Release\Veritec.Dynamics.CI.PowerShell.dll"
 
#### 1. Office 365 Method for logging in
$currentUser = Get-Credential -UserName "youruser@yourorg.onmicrosoft.com" -Message "D365 Credentials:"
$userName = $currentUser.UserName
$password = $currentUser.GetNetworkCredential().Password

$connectString =  "AuthType=Office365;Url=https://yourorg.crm6.dynamics.com;RequireNewInstance=True;UserName=$userName;Password=$password"

#### 2. S2S method Method for logging in (Note: AppId requires an app registration in Azure AD.)
#$connectString = "AuthType=OAuth;Url=https://yourorg.crm6.dynamics.com;AppId=yourAppIDGuid;UserName=youruser@yourorg.onmicrosoft.com;RedirectUri=https://yourorg.crm6.dynamics.com;LoginPrompt=Always;TokenCacheStorePath=c:\temp\mytoken;RequireNewInstance=True;"
 
# Export Dynamics Solutions from source
Export-DynamicsSolution `
    -ConnectionString $connectString `
    -SolutionName "Solution1;Solution2;" `
    -SolutionDir "..\LocalSolutionDirectory" `
 
# Export Dynamics Data Pre Solution
Export-DynamicsData `
      -ConnectionString $connectString `
      -FetchXMLFile ".\FetchXMLQueriesPre.xml" `
      -OutputDataPath ".\SourceDataPre"    
 
# Export Dynamics Data Post Solution
Export-DynamicsData `
      -ConnectionString $connectString `
      -FetchXMLFile ".\FetchXMLQueriesPost.xml" `
      -OutputDataPath ".\SourceDataPost"    
```
Example PowerShell to import the above exported Dynamics Solution and Reference Data
```powershell
# Reference PowerShell module
import-module "C:\VSRC\Dynamics 365 Practice\Veritec.Dynamics.CI\Veritec.Dynamics.CI.PowerShell\bin\Release\Veritec.Dynamics.CI.PowerShell.dll"
 
#### 1. Office 365 Method for logging in
$currentUser = Get-Credential -UserName "youruser@yourorg.onmicrosoft.com" -Message "D365 Credentials:"
$userName = $currentUser.UserName
$password = $currentUser.GetNetworkCredential().Password

$connectString =  "AuthType=Office365;Url=https://yourorg.crm6.dynamics.com;RequireNewInstance=True;UserName=$userName;Password=$password"

#### 2. S2S method Method for logging in (Note: AppId requires an app registration in Azure AD.)
#$connectString = "AuthType=OAuth;Url=https://yourorg.crm6.dynamics.com;AppId=yourAppIDGuid;UserName=youruser@yourorg.onmicrosoft.com;RedirectUri=https://yourorg.crm6.dynamics.com;LoginPrompt=Always;TokenCacheStorePath=c:\temp\mytoken"
 
# Import Dynamics Data Pre
Import-DynamicsData `
    -ConnectionString $connectString `
    -TransformFile "transforms.json" `
    -InputDataPath ".\SourceDataPre"
 
# Import Dynamics Solutions to target
Import-DynamicsSolution `
          -ConnectionString $connectString `
          -SolutionName "Solution1;Solution2" `
          -SolutionDir "..\LocalSolutionDirectory" `

# Import Dynamics Data Post
Import-DynamicsData `
    -ConnectionString $connectString `
    -TransformFile "transforms.json" `
    -InputDataPath ".\SourceDataPost"
```
# Additional Items
##  Set Plugin Status
The following cmdlet let's you enable or disable a plugin. The PluginStepNames is a semicolon delimited list of plugin names from the "SDK Message Processing Steps" section of the solution manager. Use $true or $false with the setEnabled parameter to enable/disable the plugins.
```powershell
Set-PluginStatus `
    -ConnectionString $connectString `
    -PluginStepNames 'PluginName1;PluginName2' `
    -setEnabled $true 
```

## Set Auto Number
Simple cmdlet to set the autonumber seed value. To keep this as a safe operation, the cmdlet will first ensure that there is no data present for the given entity. You can override this behaviour with the 'Force=$true' parameter
```powershell
Set-AutoNumberSeed `
    -ConnectionString $connectString `
    -EntityName 'example_entity' `
    -AttributeName 'example_attribute' `
    -Value '20050000'
```

## Transform File Constants
The {DESTINATION-ROOT-BU} constant can be used within your transform file so save having to look up the GUID manually. An example of an entry is:
```json
[
  {
    TargetEntity: "businessunit",
    TargetAttribute : "parentbusinessunitid",
    TargetValue : "*",
    ReplacementValue: "{DESTINATION-ROOT-BU}"
   },
   {
     TargetEntity: "businessunit",
     TargetAttribute: "businessunitid",
     TargetValue: "9CDB1F7D-2F2A-E811-A853-000D3AD07676",
     ReplacementValue : "{DESTINATION-ROOT-BU}"
   }
]
```

## Replacement Attribute
This feature gives you the ability to target a certain attribute and value to be replaced, but then replace a different attribute with a value. For example, in the below scenario you can see this being used for a custom config entity where a row with an attribute sol_name="DebugToolBar" should have the sol_valuestring attribute replaced with TRUE
```json
[
  {
    "TargetEntity": "sol_configurationsetting",
    "TargetAttribute": "sol_name",
    "TargetValue": "DebugToolbar",
    "ReplacementAttribute": "sol_valuestring",
    "ReplacementValue": "TRUE"
  }
]
```