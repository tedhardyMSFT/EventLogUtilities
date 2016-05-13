<#
.SYNOPSIS
Lists all messages for registered legacy event sources on the local machine

.DESCRIPTION
Iterates through the local registry for all registered legacy event sources. For each
event source, retrieve the message DLL(s) associated with that event source.

For all eventSource and message DLL combinations, get the messages that the event source
can reference for events, attempt to categorize into Message Types based upon Message ID and 
available event message documentation.
(Note:Event developers do not always follow the documentation - message type may be inaccurate.)

Write all the output to a Tab separated file, which can be imported into Excel.

Messages with Tabs, newlines will have those characters replaced with tokens:
	Tab => [TAB]
	Newline => [NL]

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


## path to msgdump.exe application - used to extract dll message exports
$MsgDumpLocation = "$env:userprofile\tools\msgdump.exe"

## Output file location
$DestOutputPath = "$env:userprofile\LegacyEventSourceMessages-$OSCaption-$OSVersion.TSV"

# test if required exe is in expected location
if ((Test-Path $MsgDumpLocation) -eq $false)
{
    write-host -ForegroundColor Red "Required app: MsgDump.exe does not exist at the location: {$MsgDumpLocation}"
    write-host -ForegroundColor Red "Please update the script with the correct location of MsgDump.exe and run it again."
    exit;
}

Write-Host -ForegroundColor Green "Writing output to $DestOutputPath";

# reference: http://msdn.microsoft.com/en-us/library/aa363661(v=vs.85).aspx
# to remove a type of file (e.g., you don't want Event Messages), just remove EventMessageFile from the list.
[String]$MessageFileNames = @("EventMessageFile", "ParameterMessageFile", "CategoryMessageFile")

# root of where legacy event sources are stored.
$LegacyEventLog = Get-Item "HKLM:\System\CurrentControlSet\Services\EventLog"

$LegacyChannelNames = $LegacyEventLog.GetSubKeyNames()

# array of unique source and message file paths, separated by pipe character
[String]$LegacySourceMessageFiles = @()

# dictionary of sources with source as the key, and array of message file paths
[HashTable] $LegacyMessageFiles = @{}

# field delimiter for output file
$DelimiterCharacter = "`t"

# header line for nice output
$OutputFileHeaderLine = "ChannelName`tEventSource`tExportFile`tExportMessageID`tExportMessageIDHex`tMessageId`tMessageType`tFacilityCode`tLocaleId`tMessage"

# Message type is the leading nybble of the message ID. It is not always consistently used by event source developers (it isn't mandated in the generation process)
# so the list isn't perfect, they are "more right than wrong" though.
[HashTable]$MessageTypes = @{}
# populate the message types hashtable for inserting description in output file
$MessageTypes.Add("0x0","Message Description")
$MessageTypes.Add("0x1","Keywords")
$MessageTypes.Add("0x2","Message Description")
$MessageTypes.Add("0x3","OpCode")
$MessageTypes.Add("0x4","Message Description")
$MessageTypes.Add("0x5","Event Level")
$MessageTypes.Add("0x6","Message Description")
$MessageTypes.Add("0x7","Task")
$MessageTypes.Add("0x8","Message Description")
$MessageTypes.Add("0x9","Log Link")
$MessageTypes.Add("0xa","Message Description")
$MessageTypes.Add("0xb","Message Description")
$MessageTypes.Add("0xc","Message Description")
$MessageTypes.Add("0xd","Resource String")
$MessageTypes.Add("0xe","Message Description")
$MessageTypes.Add("0xf","Message Description")

write-host "Enumerating all message files for LEGACY event sources"

#Loop over all legacy channel names (Application, Security, System, Hardware, etc...) that exist on the system.
# note: To read the security event sources, this script must be run from an elevated prompt.
foreach ($ChannelName in $LegacyChannelNames)
{
    $LegacySourceRegkey = $LegacyEventLog.OpenSubKey($ChannelName)

    #Get the list of event sources registered for that event channel
    $LegacySourceList = $LegacySourceRegkey.GetSubKeyNames()

    # each legacy channel can/will have multiple event sources registered.
    foreach ($LegacySourceName in $LegacySourceList)
    {
        # combine channel name and event source into a single string
        # this enables the channel name to be carried to the next phase (parsing)
        $ChannelSource = $ChannelName + ":" + $LegacySourceName
        #get the registry sub-key for the legacy event source name
   
        $LegacySource = $LegacySourceRegkey.OpenSubKey($LegacySourceName)

        # each source can have zero, one, or more "message files" 
        # whether for event, parameter, or category.
        $ValueNameList = $LegacySource.GetValueNames()

        # check each of the source registry keys value names
        foreach ($ValueName in $ValueNameList)
        {
            #check refrence file names for a value
            #if ($ValueName.Contains("MessageFile") -eq $true)
            if ($MessageFileNames.Contains($ValueName) -eq $true)
            {
                # found a match, get the path (normalize to lowercase)
                # this could be a single file or multiple files separated by semi-colon.
                $ParameterFile = $LegacySource.GetValue($ValueName).ToString().ToLower().Trim()

                # check if it contains multiple file paths.
                if($ParameterFile.Contains(';') -eq $true)
                {
                    #split multi value parameter file entries, and test each one for an event source.
                    $MultiParameterFile = $ParameterFile.Split(';')

                    foreach ($SingleParameterFile in $MultiParameterFile)
                    {
                       # check that the message file exists
                        $FileExists = Test-Path $SingleParameterFile

                        if ($FileExists -eq $true)
                        {
                            # file exists, create an array string (concatenated Source and Parameter file path, separated by pipe.)
                            $SourceMessageFile = $ChannelSource + "|" + $SingleParameterFile

                            # Check if array has the Source + message file value.
                            if ($LegacySourceMessageFiles.Contains($SourceMessageFile) -eq $false)
                            {
                                # nope, add it.
                                $LegacySourceMessageFiles = $LegacySourceMessageFiles + $SourceMessageFile
                            }

                            # check if the source exists in the Hashtable (at all)
                            if ($LegacyMessageFiles.ContainsKey($ChannelSource))
                            {
                                # yes, only add the message file if it doesn't already exist.
                                # get the array of message files stored for the event source
                                [string]$Sourcefiles = $LegacyMessageFiles[$ChannelSource]

                                #if the message files for that source isn't a match add it.
                                if ($Sourcefiles.Contains($SingleParameterFile) -eq $false)
                                {
                                    # add the new message file to the array in the hashtable.
                                    $LegacyMessageFiles[$ChannelSource] = $LegacyMessageFiles[$ChannelSource] + $SingleParameterFile
                                }
                            }
                            else
                            {
                                # create the array to be stored in the hashtable
                                [array]$newArray = ( $SingleParameterFile )

                                Write-Host "Adding $ChannelSource to dictionary";
                                # add the source name and the array in the hashtable.
                                $LegacyMessageFiles.Add($ChannelSource, $newArray)
                            }
                        }
                        else
                        {
                            #file doesn't exist - write an error message.
                            write-host -ForegroundColor Yellow "source $ChannelSource Message File: $SingleParameterFile does not exist!"
                        }
                    } # foreach ($SingleParameterFile in $MultiParameterFile)
                }
                else
                {
                    # single parameter file path scenario.

                    # check for empty/null entry (causes test-path to barf)
                    if (($ParameterFile -ne $null) -and ($ParameterFile.Trim().Length -ne 0 ))
                    {
                        # check that the message file exists
                        $FileExists = Test-Path $ParameterFile

                        # does file exist?
                        if ($FileExists -eq $true)
                        {
                            # file exists, create an array string (concatenated Source and Parameter file path, separated by pipe.)
                            [String] $SourceMessageFile = $ChannelSource + "|" + $ParameterFile

                            # Check if array has the Source + message file value.
                            if ($LegacySourceMessageFiles.Contains($SourceMessageFile) -eq $false)
                            {
                                #nope, add it.
                                $LegacySourceMessageFiles = $LegacySourceMessageFiles + $SourceMessageFile
                            }

                                # check if the source exists in the Hashtable (at all)
                            if ($LegacyMessageFiles.ContainsKey($ChannelSource))
                            {
                                # yes, only add the message file if it doesn't already exist.
                                # get the array of message files stored for the event source
                                [string]$Sourcefiles = $LegacyMessageFiles[$ChannelSource]

                                #if the message files for that source isn't a match add it.
                                if ($Sourcefiles.Contains($ParameterFile) -eq $false)
                                {
                                    # add the new message file to the array in the hashtable.
                                    $LegacyMessageFiles[$ChannelSource] = $LegacyMessageFiles[$ChannelSource] + $ParameterFile
                                }
                            }
                            else
                            {
                                # create the array to be stored in the hashtable
                                [array]$newArray = ( $ParameterFile )

                                Write-Host "Adding $ChannelSource to dictionary";
                                # add the channel:source name and the array in the hashtable.
                                $LegacyMessageFiles.Add($ChannelSource, $newArray)
                            }
                        }
                        else
                        {
                            # file doesn't exist, write error message
                            write-host -ForegroundColor Yellow "source $ChannelSource Message File: $ParameterFile does not exist!"
                        }
                    } ## if (($ParameterFile -ne $null) -and ($ParameterFile.Trim().Length -ne 0 ))
                } ## else if($ParameterFile.Contains(';') -eq $true)
            } ## if ($MessageFileNames.Contains($ValueName) -eq $true)
        } ## foreach ($ValueName in $ValueNameList)
    } ## foreach ($LegacySource in $LegacySourceList)
} ## foreach ($ChannelName in $LegacyChannelNames)

write-host $LegacyMessageFiles.Count "unique source sources found."

#
# End of Get event source information. Now:
# for each EventSource/MessageFile combination:
#	retrieve messagefile exports via msgDump.exe
#	parse the output
#	add to output array
# save output array to file.
#

# array of all eventsource, ID, and message values
[Array]$OutputMessages = @("")

# set first line to header values line.
$OutputMessages = $OutputMessages[0] = $OutputFileHeaderLine

# now to iterate over all Channel + event sources collected above.
foreach($EventSourceName in $LegacyMessageFiles.Keys)
{
    $ChannelEventPair = $EventSourceName.Split(':');

    $LegacyEventChannel = $ChannelEventPair[0];
    $LegacyEventSourceName = $ChannelEventPair[1];


    # get the array of message files for that event source
    $eventSourceMessageFile = $LegacyMessageFiles[$EventSourceName]

    # each eventSource can have one or more resource file (called Message file) for storing strings
    for ($MessageFileIndex = 0; $MessageFileIndex -lt $eventSourceMessageFile.Count; $MessageFileIndex++)
    {
        # for each message DLL reset the array of message start lines.
        [Array]$MessageStartLines = @()

        # reference the value once - makes the code cleaner.
        $MessageFileName = $eventSourceMessageFile[$MessageFileIndex]

        # run app to get all exported message strings from the Message File library
        $SourceMessages = . $MsgDumpLocation $MessageFileName

        write-host "For source:"$LegacyEventSourceName "getting Messages from message file:"$MessageFileName

        [Array]$MessageStartLines = @()

        # walk through the event source messages and look for the message header line marker.
        for($LineNumber = 0; $LineNumber -lt $SourceMessages.Count; $LineNumber++)
        {
            $IdString = ($SourceMessages[$LineNumber]).StartsWith("ID 0x")
            if ($IdString -eq $true)
            {
             #write-host "($LineNumber) "$SourceMessages[$LineNumber]
             $MessageStartLines = $MessageStartLines + $LineNumber
            }
        } # for($LineNumber = 0; $LineNumber -lt $SourceMessages.Count; $LineNumber++)

        if($MessageStartLines.Count -eq 0)
        {
            write-host -ForegroundColor Yellow "Skipping message file that contains no messages:$MessageFileName"
        }
        else
        {
            # loop through the messages lines in the array to identify the number of messages and on what line they start.
            # the first line is a header line that contains ID (in int and Hex) and language code
            for($i = 0; $i -lt ($MessageStartLines.Count-1); $i++)
            {
                # the next array entry is where the next message starts
                $MessageEndLine = $i+1
                #write-host "Message starting on line "$MessageStartLines[$i]", ending "($MessageStartLines[$MessageEndLine] -1)

                $SingleLineMessage = ""

                # example message header line looks like this: 
                #  ID 0x30000000 (805306368) Language: 0409
                #     

                $HeaderLine = $SourceMessages[$MessageStartLines[$i]]

                #write-host $HeaderLine

                # split it into an array by each spaces.
                $HeaderParts = $HeaderLine.Split(" ")

                $MessageId = $HeaderParts[2].Replace('(','').Replace(')','')
                $MessageIdHex = $HeaderParts[1]
                $MessageLanguageCode = $HeaderParts[4]
                # get the first three characters, which are the Message Description type ID
                $MessageDescriptionType = $MessageIdHex.Substring(0,3).ToLower()

                # Look up the Description Type ID in the hashtable created above.
                $MessageDescription = $MessageTypes[$MessageDescriptionType]

                # get the last 4 characters of the MessageID
                $EventSourceMessageIdHex = $MessageIdHex.Substring(6,4)

                # now convert the value to decimal
                $EventSourceMessageIdDec = [Convert]::ToInt32($EventSourceMessageIdHex,16)

                # pull out the facility code (For completeness, use case is not clear)
                $MessageFacilityCode = $MessageIdHex.Substring(3,3)

                # build an output line using the fields from the message header
                $SingleLineMessage = $LegacyEventChannel + $DelimiterCharacter + $LegacyEventSourceName + $DelimiterCharacter + $MessageFileName + $DelimiterCharacter + $MessageId + $DelimiterCharacter + $MessageIdHex + $DelimiterCharacter + $EventSourceMessageIdDec + $DelimiterCharacter + $MessageDescription  + $DelimiterCharacter + $MessageFacilityCode + $DelimiterCharacter + $MessageLanguageCode + $DelimiterCharacter

                #write-host "Message ID" $MessageId $MessageIdHex "Language Code" $MessageLanguageCode

                # skip first message line because it doesn't contain message text. It must be parsed separately.
                # concatenate the messages lines together, removing eventlog newline and tab replacement characters.
                for($MessageLine = $MessageStartLines[$i]+1; $MessageLine -lt $MessageStartLines[$MessageEndLine]; $MessageLine++)
                {
                    # replacing a subset of formatting characters used in event logs, see: http://msdn.microsoft.com/en-us/library/windows/desktop/ms679351(v=vs.85).aspx
                    $SingleLineMessage = $SingleLineMessage + $SourceMessages[$MessageLine].Trim().Replace("%n","[NL]").Replace("%t","[TAB]").Replace("%b"," ").Replace("%0"," ").Replace("\n","[NL]").Replace("\t","[TAB]")
                }

				## append message to output
                $OutputMessages = $OutputMessages + $SingleLineMessage
            } # for($i = 0; $i -lt ($MessageStartLines.Count-1); $i++)
        } # if($MessageStartLines.Count -eq 0)
    } # for ($MessageFileIndex = 0; $MessageFileIndex -lt $eventSourceMessageFile.Count; $MessageFileIndex++)
} # foreach($EventSourceName in $LegacyMessageFiles.Keys)

## now that all message files for an event source are listed, write them to a single file.
Write-Host -ForegroundColor Green "Writing legacy event source messages to file: $OutputMessages"

$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
[System.IO.File]::WriteAllLines($DestOutputPath, $OutputMessages, $Utf8NoBomEncoding)

