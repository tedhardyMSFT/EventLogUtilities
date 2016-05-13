## EventLogUtilities
# What's here?
# Short answer:
Two scripts for retrieving event data from the local system.

# Long answer:
Get-LegacyEventSourceMessages.ps1 - retrieves event information (event message, Event ID, event log name, etc...) for legacy Windows event log event sources. (Legacy meaning pre-Windows Vista - as of Windows Vista the event log was re-written)

Get-ProviderEventMessages.ps1 - retrieves event information (eventID, keywords, opcodes, event channel, event message, and event template) for Event Providers (i.e., an event manifest has been installed on the local system defining the events, channels, descriptions, xml templates) for all event providers on the system.

# What do they generate?

Both output Tab Separated Value files that can be imported into Excel or Hadoop or just read by any text editor.

# Why is this useful?

The output can be used to view all possible events that can be logged on the local system and their messages to determine if they apply to operations, security, or compliance purposes.
