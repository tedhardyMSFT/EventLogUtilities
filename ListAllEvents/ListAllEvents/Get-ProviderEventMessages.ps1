<#
.SYNOPSIS
Lists all event message data for registered manifested event providers on the local machine.

.DESCRIPTION
Iterates through the registered providers for all provider data. For each
provider, retrieve the EventIDs/Versions, Keywords, Destinaion Channels, Message Descriptions and Templates 
associated with that event provider.

Write all the output to a Tab separated file, which can be imported into Excel.

Since event providers include Event Tracing for Windows (ETW) events, if no destination channel is specified then
those events have "ETW" as the destination channel. Some ETW events have a destination channel, usually an
Analytic or Debug channel. Events cannot be determined if they are ETW or Event Log events - that is determined by
the output channel type.

Limitation:
	For events with Custom Templates (events have <UserData> and not <EventData> event payloads)
	the event XML template does not include the custom template information.


Messages with Tabs, newlines will have those characters replaced with spaces.
This preserves a single-line per event message format, making Excel and other parsing possible.

.OUTPUTS
Writes a Tab-Delimited file.

.INPUTS
No parameters or inputs used.

.NOTES
	Written by Ted Hardy
#>

# get major/minor/build version of os for versioned output
$OSVersion = (get-wmiobject -class Win32_OperatingSystem).Version;
# get caption for OS
$OSCaption = (get-wmiobject -class Win32_OperatingSystem).Caption.Replace(" ","");

# output path location
$DestOutputPath = "$env:userprofile\ProviderMessages-$OSCaption-$OSVersion.TSV"

Write-Host -ForegroundColor Green "Writing output to $DestOutputPath";

# get the ball rolling all providers.
# if interested in a single provider replace the * with the provider name
$Providers = Get-WinEvent -ListProvider *

Write-Host -ForegroundColor Green "Retrieved $($Providers.count) providers from the local system.";

# array for output
$ProviderEvents = @();

# build header line for columns
# $EventData = "$Prov.ProviderName`t$EventID`t$DestinationLog`t$EventLevel`t$keywords`t$Tasks`t$OpCodes`t`"$EventDescription`"`t$EventTemplate"
$ProviderEvents = $ProviderEvents + "ProviderName`tEventId`tEventVersion`tEventChannel`tEventLevel`tKeywords`tTasks`tOpCodes`tEventDescriptionText`tEventXmlTemplate"

# iterate over providers
foreach ($Prov in $Providers)
{
    if ($Prov.Events.Count -eq 0)
    {
        write-host -ForegroundColor Yellow "Skipping legacy event source:$($Prov.Name)";
    } # if if ($Prov.Events.Count -eq 0)
    else
    {
        Write-Host -ForegroundColor Green "Processing:$($Prov.Name) with:$($Prov.Events.Count) events";

        # iterate over all events in a given provider
        foreach ($event in $prov.Events)
        {
            $EventID = $event.Id;
            $EventVersion = [String]::Empty;
            $DestinationLog = $event.LogLink.LogName;
            $isSystemLog = $event.LogLink.IsImported;
            $OpCodes = [String]::Empty;
            $keywords = [String]::Empty;
            $Tasks = [String]::Empty;
            $EventLevel = [String]::Empty;
            $KeywordValues = [String]::Empty;
            $EventDescription = $event.Description;
            $EventTemplate = $event.Template
    
            #todo: Add Event Version (which is an optional field. ETW events don't have versions.)

            if (($DestinationLog -eq $null) -or ($DestinationLog -eq ''))
            {
                $DestinationLog = 'ETW';
            }

            if ($event.Opcode.Count -gt 0)
            {
                foreach($OpCode in $event.Opcode)
                {
                    $OpCodes += $opcode.DisplayName + ';';
                }
            }

            if ($OpCodes -eq ';')
            {
                $OpCodes = 'None'
            }


            if ($event.Keywords.Count -eq 0)
            {
                # no keywords
                $Keywords = 'None'
            }
            else
            {
                # enum all keyword values
                foreach ($keyword in $Event.Keywords)
                {
                    if ($keyword.Name.Length -gt 0)
                    {
                        $keywords += $keyword.Name.Trim() + ';'
                        $keywordValue += $KeywordValues + ';'
                    }
                } # foreach ($keyword in $Event.Keywords)
            }

            if ($keywords -eq ';')
            {
                $keywords = 'None'
            }

            if($event.Task.length -gt 0)
            {
                foreach ($Task in $event.task)
                {
                    $Tasks += $Task.Name + ';'
                }
            }
    
            if ($Tasks -eq ';')
            {
                $Tasks = 'None'
            }

            if ($event.Version -eq $null)
            {
                $EventVersion = ''
            }
            else
            {
                $EventVersion = $event.Version
            }

            if ($event.Level -eq $null)
            {
                $EventLevel = 'UnDefined'
            }
            else
            {
                $EventLevel = $Event.Level.Name
            }



            if ($EventDescription.Length -gt 0)
            {
                # remove CR/LF from template so it is on a single line
                $EventDescription = $EventDescription.Replace("`n",' ');

                $EventDescription = $EventDescription.Replace("`r",' ');
                
                # remove tab characters from template to not interfere with field delimiters
                $EventDescription = $EventDescription.Replace("`t",' ');

            }


            if ($EventTemplate.Length -gt 0)
            {
                # some providers just put CR character, not combined.
                $EventTemplate = $EventTemplate.Replace("`r",' ');
                # remove remaining LF characters
                $EventTemplate = $EventTemplate.Replace("`n",' ');
                
                # remove tab characters from template to not interfere with field delimiters
                $EventTemplate = $EventTemplate.Replace("`t",' ');

            }

            
            if ($EventDescription -eq [String]::Empty)
            {
                $EventDescription = "No Description";
            }

            if ($EventTemplate -eq [String]::Empty)
            {
                $EventTemplate = "No Event Template";
            }
            

            # assemble the event data
            $EventData = $Prov.Name.ToString()+"`t$EventID`t$EventVersion`t$DestinationLog`t$EventLevel`t$keywords`t$Tasks`t$OpCodes`t`"$EventDescription`"`t$EventTemplate"

            # add event information to output array
            $ProviderEvents = $ProviderEvents + $EventData
        } # foreach ($event in $prov.Events)
    } # else if ($Prov.Events.Count -eq 0)
} # foreach ($Prov in $Providers)


Write-Host -ForegroundColor Green "Writing output to $DestOutputPath";
# don't emit the UTF-8 BOM encoding, it causes Cosmos to act weirdly.
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
# write all lines in output array to text file.
[System.IO.File]::WriteAllLines($DestOutputPath, $ProviderEvents, $Utf8NoBomEncoding)

