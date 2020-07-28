# OutlookHeaders Add-in
*Microsoft Office Outlook add-in for modifying mail headers &amp; outbound emails.*

### Description

OutlookHeaders is a simple Microsoft Outlook add-in for Windows that allows to adjust outbound email messages.

Overall this add-in can do the following:

- Add arbitrary http headers to your outbound email messages.
- Overwrite mail client name, so that your email doesn't appear to have come from Microsoft Outlook.
- Suppress initial `received:` header, to prevent inclusion of the public IP address where the email was composed. (Outlook by default adds your public IP to every message.)
- Adjust the format of your replies and forwarded emails to plaintext, or HTML, regardless of the format of the original message.
- At the same time, if adjusting the format of an email to HTML, this add-in allows to adjust the Cascading Style Sheets (or CSS) for it.
- This add-in also allows to automatically remove the tracking pixel from your replies and forwarded emails for specific domains.
- All options listed above can be applied per individual email account, or for all accounts in Outlook.
- This add-in supports automated installation, which allows to set it up from a pre-configured file.
- This add-in supports non-interactive installation in an Active Directory environment, or on a corporate network via Group Policy Objects (or GPOs.)

### Screenshot

![Alt text](https://dennisbabkin.com/php/imgs2/scrsht_olh_05.png "'Customize Outbound Mail' window")

### Manual

To learn more about this add-in, check its [official manual](https://dennisbabkin.com/php/docs.php?what=olh).

### Release Build

If you don't want to build and code-sign this add-in yourself, you can download the latest [release build here](https://dennisbabkin.com/olh).

### Prerequisites

- Microsoft Outlook 2007 for Windows, or later
- [Microsoft .NET Framework 3.5](https://www.microsoft.com/en-us/download/details.aspx?id=21)
- [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=56961)

### Build Instructions

The original intentions of making this add-in was to make it backward compatible with older versions of Microsoft Office. (The earliest I chose was Office 2007.) That is why to build this add-in I had to use older versions of the Visual Studio and the .NET Framework.

There are four Visual Studio solutions that are involved in building this add-in. Each has to be built in this particular order:

1. **OutlookHeaders** Solution (`OutlookHeaders.sln` file) was originally designed and built in `Visual Studio 2008, SP1 v.9.0`

   This solution contains the `OutlookHeaders` project with the actual code for this add-in. It's written in C#. This project requires the following references to be able to work with the Office components:
   
    - `Microsoft.Office.Interop.Outlook`
    - `Microsoft.Office.Tools.Common.v9.0`
    - `Microsoft.Office.Tools.Outlook.v9.0`
    - `Microsoft.Office.Tools.v9.0`
    - `Microsoft.VisualStudio.Tools.Applications.Runtime.v9.0`
    - `Office`
    - `System.AddIn`
    
   The `OutlookHeaders` project in this solution had the `Post-Build` event that would sign the resulting `OutlookHeaders.dll`, as the `$(TargetPath)` file, with our code-signing certificate. Thus, to be able to load it into the Outlook without seeing a security warning, you will need to provide your own code-signing certificate to digitally sign it. Without digitally signing your add-in Microsoft Office may refuse to load your add-in. For organizations, I believe there is a [way to self-sign your office add-ins](https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-sign-office-solutions?view=vs-2019).

2. **CustomActionOutlookHeadersInstaller** Solution (`CustomActionOutlookHeadersInstaller.sln` file) was originally designed and built in `Visual Studio 2010 v.10.0`

   This solution contains the `CustomActionOutlookHeadersInstaller` project with the code for the [custom actions](https://docs.microsoft.com/en-us/windows/win32/msi/custom-actions) for the MSI installer. It is written in C#. The project requires the following references to interface with the WiX installer:
   
    - `Microsoft.Deployment.WindowsInstaller` (included in the [WiX Toolset](https://wixtoolset.org/) library)
   
   The `CustomActionOutlookHeadersInstaller` project in this solution had the `Post-Build` event that would sign the resulting `CustomActionOutlookHeadersInstaller.CA.dll`, as the `$(TargetDir)$(TargetName).CA$(TargetExt)` file, with our code-signing certificate. This signature was provided only to establish end-user trust and is not required for this add-in to be loaded into Microsoft Office.
   
3. **OutlookHeadersInstaller** Soltution (`OutlookHeadersInstaller.sln` file) was originally designed and built in `Visual Studio 2010 v.10.0` as a `Setup Project` built with the use of the [WiX Toolset](https://wixtoolset.org/) library v.3.7.

   This solution contains the `OutlookHeadersInstaller` project with the `Windows Installer XML` markup to build the main MSI installer for this add-in, that provides code to install, uninstall, repair and upgrade it on client workstations. The project requires the following references to the WiX library components:
   
    - `WixNetFxExtension`
    - `WixUtilExtension`

   The `OutlookHeadersInstaller` project in this solution had the `Post-Build` event that would sign the resulting `OutlookHeadersInstaller.msi`, as the `$(TargetPath)` file, with our code-signing certificate. (Only SHA-1 signature was used to provide compatibility with MSI signing limitations.) This signature was provided only to establish end-user trust and is not required for this add-in to be loaded into Microsoft Office.

4. **BootstrapperOutlookHeaders** Soltution (`BootstrapperOutlookHeaders.sln` file) was originally designed and built in `Visual Studio 2010 v.10.0` as a `Bootstrapper Project`, built with the use of the [WiX Toolset](https://wixtoolset.org/) library v.3.7.

   This solution contains the `BootstrapperOutlookHeaders` project with the XML markup to build the Bootstrapper installer for this add-in to be able to install needed prerequisite components (listed above). The project requires the following references to the WiX library components:
   
    - `WixBalExtension`
    - `WixNetFxExtension`
    - `WixUtilExtension`
    
   The `BootstrapperOutlookHeaders` project in this solution had the `Post-Build` event that would de-construct and sign components of the resulting bootstrapped installer. This was necessary to pass the required bootstrapper authentication of the `engine.exe` and is not necessary for this add-in to be loaded into Microsoft Office. The following commands were provided:
   
       del "%tmp%\engine.exe"
       "C:\Program Files (x86)\WiX Toolset v3.7\bin\insignia.exe" -ib "$(TargetFileName)" -o "%tmp%\engine.exe"
       command-to-code-sign "%tmp%\engine.exe"
       "C:\Program Files (x86)\WiX Toolset v3.7\bin\insignia.exe" -ab "%tmp%\engine.exe" "$(TargetFileName)" -o "$(TargetFileName)"
       command-to-code-sign "$(TargetPath)"
       del "%tmp%\engine.exe"

    where `command-to-code-sign` must be replaced with your own command to digitally code-sign that specific file.
    
### Installation

The reason we're building several versions of the installer is to provide for different ways of application of this add-in:

- The bootstrapper installer (labeled for end-users as "Bundled Installer") is an easy way to install the add-in with a possibility to download and install all needed prerequisited in a single UI package.

- The standalone MSI installer is a light-weight version, that was designed to support non-interactive installation on multiple workstations, such as Microsoft Active Directory, via Group Policy Objects. But it requires for all prerequisite components (listed above) to be already present on the workstations.

[Check here](https://dennisbabkin.com/php/docs.php?what=olh&ver=1.0.2#installation) for details of the installation.


--------------


Submit suggestions & bug reports [here](https://www.dennisbabkin.com/sfb/?what=bug&name=OutlookHeaders+Add-in&ver=Github).
