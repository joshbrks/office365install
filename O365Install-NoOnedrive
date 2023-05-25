if (Test-Path -Path 'C:\Support') {
} else {
    New-Item -Path 'C:\Support\' -ItemType Directory
}

#Copy Installer
Invoke-WebRequest -uri "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365ProPlusRetail&platform=Def&language=en-us&TaxRegion=pr&correlationId=3dda7fc1-09ce-43fe-a9e3-8e6b7ea16401&token=0642b10e-77b5-44d4-a9df-982c54186524&version=O16GA&source=O15OLSO365&Br=2" -OutFile C:\Support\OfficeSetup.exe

New-Item C:\Support\OfficeConfig.xml -ItemType File
Set-Content C:\support\OfficeConfig.xml @"
<Configuration ID="6de936b1-0302-4bba-8810-b647b8ed37c7" DeploymentConfigurationID="00000000-0000-0000-0000-000000000000">
  <Add OfficeClientEdition="64" Channel="Monthly" ForceUpgrade="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="OneDrive" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
Set-Location C:\Support\
Start-Process .\OfficeSetup.exe -Wait -ArgumentList '/configure OfficeConfig.xml'