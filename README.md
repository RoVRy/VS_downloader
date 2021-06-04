# VS_downloader
PowerShell script for downloading/updating actual Visual Studio offline package

This script:
- allows you to download to disk or upgrade to the latest standalone version of Visual Studio;
- allows you to download any edition: Community, Professional or Enterprise;
- you can select all 14 Visual Studio languages in any combination;
- the script message language is automatically selected depending on the locale of the operating system.

The script downloads only two files from the Internet: the web page of Microsoft's official site, where it finds the link and downloads the official Visual Studio online-installer. Further downloading of the packages is performed solely by the Visual Studio installer. After the script finishes, these two temporary files will be deleted.
