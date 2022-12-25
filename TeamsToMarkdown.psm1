
$Script:ProgressPreference = 'SilentlyContinue'

$Script:TeamsEnvironment = 'Global'

$Script:TeamsApiVersion  = 'Beta'

$Script:NewLine = [System.Environment]::NewLine

$Script:TeamsDefaultDelegatedPermissions = @(
'User.Read'
'User.ReadBasic.All'
'Team.ReadBasic.All'
'Channel.ReadBasic.All'
'Chat.Read'
'ChatMessage.Read'
'Files.Read.All'
'Sites.Read.All'
)

if ( -not ( Get-Module 'MarkdownPrince' ) )
{
  Import-Module 'MarkdownPrince' -Force -ErrorAction 'Stop'
}

if ( -not ( Get-Module 'Microsoft.Graph.Authentication' ) )
{
  Import-Module 'Microsoft.Graph.Authentication' -Force -ErrorAction 'Stop'
}

#  change target API version
Select-MgProfile -Name $Script:TeamsApiVersion

if ( -not ( Get-Module 'Microsoft.Graph.Users' ) )
{
  Import-Module 'Microsoft.Graph.Users' -Force -ErrorAction 'Stop'
}
  
if ( -not ( Get-Module 'Microsoft.Graph.Teams' ) )
{
  Import-Module 'Microsoft.Graph.Teams' -Force -ErrorAction 'Stop'
}

$Script:MarkdownArgs = @{
  Content                  = ''
  UnknownTags              = 'Bypass'
  GithubFlavored           = $true
  RemoveComments           = $true
  SmartHrefHandling        = $true
  DefaultCodeBlockLanguage = ''
  Format                   = $false
}


Function Connect-Teams
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions
  )

  if ( 
        ( -not $Script:TeamsCurrentSession ) -or 
        (      $Script:TeamsCurrentSession -and 
               ( 
                 ( $Script:TeamsCurrentSession.TenantId -ne $TenantId ) -or 
                 ( $Script:TeamsCurrentEnvironment      -ne $TeamsEnvironment     ) -or
                 ( Compare-Object -ReferenceObject $Script:TeamsCurrentSession.Scopes -DifferenceObject $TeamsDelegatedPermissions )
               ) 
        )
  )
  {
    
    #  connect using interactive authentication
    Connect-MgGraph -Environment $TeamsEnvironment -TenantId $TenantId -Scopes $TeamsDelegatedPermissions -ErrorAction 'Stop'
  
    $Script:TeamsCurrentSession     = Get-MgContext -ErrorAction 'Stop'
    
    $Script:TeamsCurrentEnvironment = $TeamsEnvironment
    
    #  get ms graph endpoint for environment
    $Script:TeamsEndPoint           = ( Get-MgEnvironment | Where-Object { $_.Name -eq $TeamsEnvironment } ).GraphEndpoint + '/' + $Script:TeamsApiVersion
    
  }  
  
}


Function Get-TeamsChat
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions,
                           [ string[] ] $TeamsChatId,
                           [ string[] ] $TeamsChatType,  #  ( 'meeting', 'group', 'oneOnOne' )
                             [ string ] $TeamsChatTopic
  )
  
  Connect-Teams -TeamsEnvironment $TeamsEnvironment -TenantId $TenantId -TeamsDelegatedPermissions $TeamsDelegatedPermissions
  
  #  get chat list ordered by last activity time desc
  Get-MgChat -All -Sort 'lastMessagePreview/createdDateTime desc'
  | Where-Object { 
    ( ( -not ( $TeamsChatId    ) ) -or ( $TeamsChatId    -and ( $_.Id       -in    $TeamsChatId    ) ) ) -and
    ( ( -not ( $TeamsChatType  ) ) -or ( $TeamsChatType  -and ( $_.ChatType -in    $TeamsChatType  ) ) ) -and
    ( ( -not ( $TeamsChatTopic ) ) -or ( $TeamsChatTopic -and ( $_.Topic    -match $TeamsChatTopic ) ) )
  }

}


Function Get-TeamsChatMember
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,    
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions,
    [ Parameter(Mandatory) ] [ string ] $TeamsChatId
  )
  
  Connect-Teams -TeamsEnvironment $TeamsEnvironment -TenantId $TenantId -TeamsDelegatedPermissions $TeamsDelegatedPermissions
  
  #  get chat member list
  Get-MgChatMember -ChatId $TeamsChatId

}

<#
 .Synopsis
  Save messages and their contents from Microsoft Teams chat to a local markdown file.

 .Description
  Save messages and their contents from Microsoft Teams chat to a local markdown file.

 .Parameter TeamsTenantId
  Microsoft Teams Tenant ID (for example : "b33cbe9f-8ebe-4f2a-912b-7e2a427f477f").

 .Parameter TeamsEnvironment
  Target Microsoft cloud name (for example : "Global").

 .Parameter TeamsDelegatedPermissions
  Delegated permissions for call Microsoft Graph REST API as a signed in user (for example : @('Team.ReadBasic.All','ChatMessage.Read') ).
  
 .Parameter TeamsChatId
  Message source Microsoft Teams internal chat id (for example : "19:b8577894a63548969c5c92bb9c80c5e1@thread.v2").
  
 .Parameter MarkdownFileName
  Markdown file name (store in DownloadDir).
  
 .Parameter DownloadDir
  Directory path for storing files downloaded from Teams (for example : "d:\t2md\download").
  
 .Parameter TrIDPathDir
  The catalog in which the program TrID (C) 2003-16 By Marco Pontello (https://mark0.net/soft-trid-e.html) is located (for example : "d:\teams_to_zulip\trid").

 .Parameter ShowProgress
  Displays a progress bar in a PowerShell command window.

 .Example
  $ConvertArgs = @{
    TeamsTenantId        = 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f'
    TeamsChatId          = '19:b8577894a63548969c5c92bb9c80c5e1@thread.v2'
    MarkdownFileName     = 'teams_chat.md'
    DownloadDir          = 'd:\t2md\download'
    TrIDPathDir          = 'd:\t2md\trid'
    ShowProgress         = $true
  }  
  ConvertFrom-TeamsChatToHtmlFile @ConvertArgs

#>
function ConvertFrom-TeamsChatToMarkdownFile
{
  [ CmdletBinding( PositionalBinding = $false ) ]
  Param
  (
    [ Parameter(Mandatory) ] [ string ] $TeamsTenantId,
                             [ string ] $TeamsEnvironment = $Script:TeamsEnvironment,    
                           [ string[] ] $TeamsDelegatedPermissions = $Script:TeamsDefaultDelegatedPermissions,
    [ Parameter(Mandatory) ] [ string ] $TeamsChatId,
                             [ string ] $MarkdownFileName,
    [ Parameter(Mandatory) ] [ string ] $DownloadDir,
    [ Parameter(Mandatory) ] [ string ] $TrIDPathDir,
                             [ switch ] $ShowProgress
  )
  
  Write-Verbose -Message "Starting: `n$($MyInvocation.MyCommand)"
  
  Connect-Teams -TeamsEnvironment $TeamsEnvironment -TenantId $TeamsTenantId -TeamsDelegatedPermissions $TeamsDelegatedPermissions
  
  #  html document object
  $HtmlDocument = New-Object -Com 'HTMLFile'
  
  try
  {
    
    Write-Verbose -Message "Get Teams user list"
    
    #  get teams user list
    $TeamsUserList = @{}
    Get-MgUser -All 
    | ForEach-Object { 
      $TeamsUserList[ $_.Id ] = @{ 
        DisplayName = $_.DisplayName 
        Mail        = $_.Mail
      } 
    }
  }
  catch
  {
    throw "Teams user list not found!"
  }  
  
  try
  {
    
    Write-Verbose -Message "Get Teams chat message list"
    
    $ExportMessageList = @{}

    #  get chat message list
    $ChatMessageList = Get-MgChatMessage -ChatId $TeamsChatId -All  
    | Where-Object { ( $_.MessageType -in ( 'message' ) ) -and ( -not ( $_.DeletedDateTime ) ) }
    | Sort-Object -Property Id
  }
  catch
  {
    throw "Teams chat message list not found!"
  }  
  
  #  get chat
  Get-TeamsChat -TenantId $TenantId -TeamsChatId $TeamsChatId  
  | Where-Object { ( $_.Id ) -and ( $_.ChatType ) }
  | ForEach-Object { 
  
    Write-Verbose -Message "Teams Chat Id : $($_.Id)"
    
    if ( -not ( $MarkdownFileName ) )  
    {
      $MarkdownFileName = $_.ChatType + '-' + ( ( [System.BitConverter]::ToString( [System.Security.Cryptography.HashAlgorithm]::Create('SHA1').ComputeHash( [System.Text.Encoding]::UTF8.GetBytes( $_.Id ) ) ) ).Replace( '-', '' ).ToLower() ) + '.md'
    }  
    
    Write-Verbose -Message "File name : $($MarkdownFileName)"
    
    #  create file
    New-Item    -Path "$($DownloadDir)\$($MarkdownFileName)" -ItemType 'File' -Force | Out-Null
    
    #  add chat title to file
    Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value ( $NewLine + '##### ' + $_.Id )
    Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value ( $NewLine + '##### ' + $_.ChatType )
    Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value ( $NewLine + '##### ' + $_.Topic )
    Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value ( $NewLine + '##### ' + $_.LastUpdatedDateTime.DateTime )
    
    #  add chat members to file
    if ( $_.ChatType -in ( 'group', 'oneOnOne' ) )  
    { 
      Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value 'Members:' 
      Get-TeamsChatMember -TenantId $TenantId -TeamsChatId $_.Id
      | ForEach-Object { 
        Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value $_.DisplayName 
      } 
    }
  }
  
  #  process message list  
  foreach ( $ChatMessage in $ChatMessageList )
  {
    
    $ExportMessageList[ $ChatMessage.Id ] = @{ 
      Id                  = $ChatMessage.Id
      CreatedDateTime     = $ChatMessage.CreatedDateTime
      FromUserId          = $ChatMessage.From.User.Id
      QuoteMark           = ''
    }
    
    Write-Verbose -Message "Teams message id: $($ChatMessage.Id)"
    
    Write-Verbose -Message "Teams HTML initial message body: `n`n$($ChatMessage.Body.Content)`n"
    
    #  replace author tag to italic
    $ChatMessageBodyContent = $ChatMessage.Body.Content
    
    $ChatMessageBodyContent = $ChatMessageBodyContent -replace '<at\s.*?>', '<strong><em>'

    $ChatMessageBodyContent = $ChatMessageBodyContent -replace '</at>'    , '</em></strong>'

    
    Write-Verbose -Message "Parse Teams HTML message body"
    
    #  parse message body html 
    $HtmlDocument.write( [System.Text.Encoding]::Unicode.GetBytes( $ChatMessageBodyContent ) )
    $HtmlDocument.Close()
    
    #  get image links
    foreach ( $Image in $HtmlDocument.Images )
    { 
    
      $ImageId  = $Image.GetAttribute( 'itemid' )
      
      $ImageUrl = $Image.GetAttribute( 'src' )
      
      $ImageItemType  = $Image.GetAttribute( 'itemtype' )
      
      if ( $ImageItemType -eq 'http://schema.skype.com/Emoji' )
      {
        #  emoji, so no need to download the image file
        $Image.SetAttribute( 'src'  , '' )
        $Image.SetAttribute( 'title', '' )
      }  
      else
      {
        #  download image file
        try
        {
          
          Write-Verbose -Message "Download image file from Teams : $($DownloadDir)\$($ImageId)"
          
          #  download image to file
          Invoke-MgGraphRequest -Uri $ImageUrl -Method 'Get' -OutputFilePath "$($DownloadDir)\$($ImageId)" -ErrorAction 'SilentlyContinue'
          
          if ( Test-Path "$($DownloadDir)\$($ImageId)" -PathType 'Leaf' )
          {
          
            #  set real file extension for image file
            & "$($TrIDPathDir)\trid.exe" @( "$($DownloadDir)\$($ImageId)", '-ae' ) > $null
            
            #  get image file full name
            $ImageFileName = ( Get-ChildItem -Path "$($DownloadDir)\$($ImageId).*" ).Name
            
            #  change image url to zulip url
            $Image.SetAttribute( 'src', "./$($ImageFileName)" )
              
          }
          else
          {
            $Image.SetAttribute( 'src', '' )
          }  
          
        }
        catch
        {
          if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
          
          Write-Verbose -Message "The file is not downloaded from Teams : $($ImageUrl) `n$($ErrorMessage)"
          
          $Image.SetAttribute( 'src', '' )
        }   

      }        
      
    }

    
    #  get final html text
    $HtmlText = $HtmlDocument.GetElementsByTagName( 'HTML' ) | Join-String -Property { $_.OuterHTML }
    
    Write-Verbose -Message "Teams HTML final message body: `n`n$($HtmlText)`n"
    
    Write-Verbose -Message "Convert Teams HTML message body to Markdown"
    
    #  convert html to markdown
    $MarkdownArgs.Content = $HtmlText
    $MarkdownText = ConvertFrom-HTMLToMarkdown @MarkdownArgs

    
    #  get attchments
    foreach ( $Attachment in $ChatMessage.Attachments )
    {
      switch ( $Attachment.ContentType )
      {
        'reference'
        {
          
          $Base64Value = [System.Convert]::ToBase64String( [Text.Encoding]::UTF8.GetBytes( $Attachment.ContentUrl ), [Base64FormattingOptions]::None )

          $EncodedUrl = 'u!' + $Base64Value.TrimEnd( '=' ).Replace( '/', '_' ).Replace( '+', '-' )    
          
          try
          {
            
            #  get drive item
            $DriveItem = Invoke-MgGraphRequest -Uri "$($Script:TeamsEndPoint)/shares/$($EncodedUrl)/driveItem" -Method 'Get' 
            
            try
            {
              
              Write-Verbose -Message "Download attachment file from Teams : $($DownloadDir)\$($DriveItem.name)"
              
              #  download drive item to file
              Invoke-WebRequest -Uri $DriveItem.'@microsoft.graph.downloadUrl' -OutFile "$($DownloadDir)\$($DriveItem.name)"
                  
              #  add link to file
              $MarkdownText = $MarkdownText + $NewLine + ( "[{0}]({1})" -f $Attachment.Name, "./$($DriveItem.name)" )
              
            }
            catch
            {
              if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
              Write-Verbose -Message "The file is not downloaded from Teams : $($Attachment.Name) `n$($ErrorMessage)"
              
              $MarkdownText = $MarkdownText + $NewLine + ( "[{0}]({1})" -f $Attachment.Name, '' )
            }
          
          }
          catch
          {
            if ( $_.ErrorDetails.Message ) { $ErrorMessage = $_.ErrorDetails.Message } else { $ErrorMessage = $_ }
            Write-Verbose -Message "The drive item is not obtained from Teams : $($Attachment.ContentUrl) `n$($ErrorMessage)"
            
            $MarkdownText = $MarkdownText + $NewLine + ( "[{0}]({1})" -f $Attachment.Name, '' )
          }  
          
        }
        'messageReference'
        {
          
          $ReferenceMessageId = ( ConvertFrom-Json -InputObject $Attachment.Content ).messageId
          
          $QuoteMark = ( $ExportMessageList[ $ReferenceMessageId ] ).QuoteMark
          
          if ( $QuoteMark ) { $QuoteMark += '>' }  else { $QuoteMark = '>' } 

          ( $ExportMessageList[ $ChatMessage.Id ] ).QuoteMark = $QuoteMark
          
          $QuotedText = ''
          foreach ( $LineText in ( $ExportMessageList[ $ReferenceMessageId ] ).MarkdownText.Split( $NewLine, [System.StringSplitOptions]::RemoveEmptyEntries ) )
          {
            $QuotedText += ( $NewLine + "$($QuoteMark)$($LineText)" )
          }  
          
          $QuotedText += ( $NewLine + $QuoteMark )
          
          #  add quote
          $MarkdownText = $QuotedText + $NewLine + $MarkdownText
          
        }        
      }
      
    }
    
    #  adapt markdown 
    $MarkdownText = $MarkdownText.Replace( '\_', '_' )
    
    $MarkdownText = $MarkdownText -replace '!\[(.*)\]\(\)', '${1}'
        
    Write-Verbose -Message "Markdown message body: `n`n$($MarkdownText)`n"
    
    #  add delimeter
    Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value '---'
    
    #  add id, author, date to message body
    $MarkdownText = 
    "##### {0}{1}**{2}** **{3}**{4}{5}" -f (
      $ChatMessage.Id,
      $NewLine,
      ( $TeamsUserList[ $ChatMessage.From.User.Id ] ).DisplayName,
      $ChatMessage.CreatedDateTime,
      $NewLine,
      $MarkdownText
    )
    
    ( $ExportMessageList[ $ChatMessage.Id ] ).Add( 'MarkdownText'  , $MarkdownText )
    
    #  add to file
    Add-Content -Path "$($DownloadDir)\$($MarkdownFileName)" -Value $MarkdownText
      
    #  display a progress bar
    if ( $ShowProgress )
    {
      $PercentComplete = [math]::Round( ( $ExportMessageList.Count / $ChatMessageList.Count ) * 100 )
      
      $ProgressPreferenceCurrent = $ProgressPreference
      
      $ProgressPreference = 'Continue'
      
      Write-Progress -Activity "File Creation in Progress" -Status "$($PercentComplete)% Complete:" -PercentComplete $PercentComplete
      
      $ProgressPreference = $ProgressPreferenceCurrent
    }  
    
  }
  
}  


Export-ModuleMember -Function Connect-Teams

Export-ModuleMember -Function Get-TeamsChat

Export-ModuleMember -Function Get-TeamsChatMember

Export-ModuleMember -Function ConvertFrom-TeamsChatToMarkdownFile
