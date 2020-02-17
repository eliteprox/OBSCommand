# OBSCommand
A Command Line Tool for Automating OBS and Title Changes in Restream.io with Microsoft Edge

Example Usage:


dotnet.exe "OBSCommand.dll" /scene="Live Feed" /startstreaming /runtime=239 /profile="ProfileName" /title="Title of the Show",true,#HashTagOrOtherText


Supported Switches and Parameter Syntax

/profile: sets the profile of OBS Studio

/scene: Switches the Current Scene

/startstreaming: Starts Streaming 

/runtime: amount of time to stream for. Format: Minutes.Seconds (Ex: 12.30)

/title: Automates Restream.io with Microsoft Edge. Must be already logged on. Format: "Title of the Show",true/false (include datestamp),#HashTagOrOtherText - the third parameter is optional.

