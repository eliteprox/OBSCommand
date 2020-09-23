# OBSCommand
A Command Line Tool for Automating OBS and Title Changes in Restream.io with Google Chrome

Example Usage:


dotnet.exe "OBSCommand.dll" /scene="Live Feed" /startstreaming /runtime=239 /profile="ProfileName" /title="Title of the Show",true,#HashTagOrOtherText


Supported Switches and Parameter Syntax

/profile= sets the profile of OBS Studio

/scene= Switches the Current Scene

/startstreaming= Starts Streaming 

/runtime= amount of time to stream for. Format: Minutes.Seconds (Ex: 12.30)

/title= Automates Restream.io with Microsoft Edge. Must be already logged on. Format: "Title of the Show"
   
/showdate For Title Only, If included the date will be appended to the title

/hashtag= For Title Only, If included will append this text to the end of the title (including /showdate)

