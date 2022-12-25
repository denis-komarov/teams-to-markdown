# TeamsToMarkdown

**TeamsToMarkdown** is a [PowerShell](https://learn.microsoft.com/powershell/) [module](https://learn.microsoft.com/en-us/previous-versions/dd901839(v=vs.85))
that that allows you to save messages and their contents from [Microsoft Teams](https://www.microsoft.com/en-us/microsoft-teams/group-chat-software) chat to a local [Markdown](https://en.wikipedia.org/wiki/Markdown) file.

**TeamsToMarkdown** has the following main functions:

- Get-TeamsChat
- Get-TeamsChatMember
- ConvertFrom-TeamsChatToMarkdownFile

## Installation

### Install the latest stable version of [PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell) (from version 7.3 and above)

### Create the necessary directories, for example:
```
d:\t2md
d:\t2md\download
d:\t2md\module
d:\t2md\module\TeamsToMarkdown
d:\t2md\trid
```

### Install the required modules from PSGallery by running a script like this in the PowerShell console:
```powershell
$PSModuleDir = 'd:\t2md\module'
$ModuleRepository = 'PSGallery'
Save-Module -Name 'Microsoft.Graph.Users' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'Microsoft.Graph.Teams' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'Microsoft.Graph.Files' -Path $PSModuleDir -Repository $ModuleRepository -Force
Save-Module -Name 'MarkdownPrince'        -Path $PSModuleDir -Repository $ModuleRepository -Force
```

### Copy the program files of [TrID](http://mark0.net/soft-trid-e.html) into the directory as follows:
Download the file [trid_w32.zip](https://mark0.net/download/trid_w32.zip) and unzip its contents into the directory so that the files appear in it:
```
d:\t2md\trid\trid.exe
d:\t2md\trid\readme.txt
```
Download the file [triddefs.zip](https://mark0.net/download/triddefs.zip) and unzip its contents into the directory so that the file appear in it:
```
d:\t2md\trid\triddefs.trd
```

### Copy the files from this repository into the directory, so that the files appear in it:
```
D:\t2md\module\TeamsToMarkdown\TeamsToMarkdown.psd1
D:\t2md\module\TeamsToMarkdown\TeamsToMarkdown.psm1
```

## Usage

### Ask your system administrator for your Microsoft Teams tenant ID. This is an identifier that looks something like this:
```
b33cbe9f-8ebe-4f2a-912b-7e2a427f477f
```

### Well, now you can start transferring messages from Teams to Zulip by running a script like this in the PowerShell console:
```powershell
$env:PSModulePath += ';' + 'd:\t2md\module'

Import-Module -Name 'd:\t2md\module\TeamsToMarkdown'

```

## 3rd party references

This module uses but does not include various external libraries and programs. Their authors have done a fantastic job.
- [Powershell SDK for Microsoft Graph](https://github.com/microsoftgraph/msgraph-sdk-powershell) - Copyright (c) Microsoft Corporation. All Rights Reserved. Licensed under the MIT license.
- [MarkdownPrince](https://github.com/EvotecIT/MarkdownPrince) - Copyright (c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.
- [TrID](http://mark0.net/soft-trid-e.html) -  Copyright (c) 2003-16 Marco Pontello. All rights reserved. Licensed as freeware for non commercial, personal, research and educational use.
