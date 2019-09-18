global newtitle := ""
global timetorun := ""
global disconnect := ""
global tagtitle := true

global newtitle := ""
global showdate := false
global hashtag := ""

for n, param in A_Args  ; For each parameter:
{
    if InStr(param, "/title=") {
        newtitle := StrReplace(param, "/title=") 
    }
    
    if InStr(param, "/showdate") {
        showdate := true
    }
    
    if InStr(param, "/hashtag=") {
        hashtag := StrReplace(param, "/hashtag=")
    }
}

try {
    wb := IEGet("https://restream.io")
    if (wb.document.title = "") {
        Failsafe()
        return
    }

    wb.Navigate("https://restream.io/titles")
    IELoad(wb)

    ;Add the date
    TimeString := ""
    if (showdate) {
        FormatTime, TimeString,, dddd MMMM d, yyyy     
        TimeString := " | " . TimeString
    }
    
    ;Add Space if needed
    if (hashtag != "") {
        hashtag := " " . hashtag
    }
    
    fulltitle := newtitle . TimeString . hashtag
    
    wb.document.getElementById("jsAllTitlesInput").value := fulltitle
    
    myScript = 
    (
    var updateAllTitlesBtn = document.getElementById("updateAllTitlesBtn");
    var inputAllTitle = $('#jsAllTitlesInput');
    
    
    var titleValue = inputAllTitle.val().trim();
    
    if (!titleValue.length) {
        $(updateAllTitlesBtn).parent().addClass('app-title_state_edit');
        inputAllTitle.addClass('input_state_error'); 
        throw new Error("Something went badly wrong!");
    }

    eventEmitter.emit('update-all-title-attempt', {});
    $(updateAllTitlesBtn).addClass("disabled");
    $('#jsChannelTitleList > .app-title').attr('class', 'app-title app-title_state_loading');
    
	var channels = $('div[data-platform][data-channel-id]');
	var total = channels.length;
	var errors = 0;
    
	$.each(channels, function (idx, channel) {
		var gameId = $(channel).find('select.form-control option:checked').val();
		var channelId = $(channel).data().channelId;

		var data = {
			channelId: channelId,
			channelTitle: titleValue,
			gameId: gameId
		};
        
        $.ajax({
			url: '/api/titles/update_channel_info',
			type: 'POST',
			dataType: 'json',
			data: data,
			success: function (data) {
				total--;
				var titleItem = $('#jsChannelTitleList > .app-title[data-channel-id="' + channelId + '"]');
				$(titleItem).removeClass('app-title_state_loading');
				if (total === 0) {
					$(updateAllTitlesBtn).removeClass('disabled');
					// Event [Update all title: success]
					eventEmitter.emit('update-all-title-success', {});
				}
			},
			error: function () {

			}
		});
    });
    
    )
    
    ;Load the necessary JavaScript using a specialized Injection Technique (Restream stopped supporting IE, but it's all we can automate easily)
    qt := chr(34)
    IID := "{332C4427-26CB-11D0-B483-00C04FD90119}" ; IID_IWebBrowserApp
    window := ComObj(9,ComObjQuery(wb,IID,IID),1)
    window.eval(myScript)
    sleep 500
    
    ;Click the button!
    wb.document.getElementById("updateAllTitlesBtn").click()
    
    IELoad(wb)
    Sleep, 500
} catch {
    Failsafe()
    return
}

Failsafe()

return
Failsafe()
{
    ;Run the OBS Studio Package
    ;qt := chr(34)
    ;script := A_ScriptDir . "/TestClient.exe " . qt . scenename . qt . " " . timetorun . " " . disconnect
    ;Run, %script%
    exit
}

IELoad(wb)    ;You need to send the IE handle to the function unless you define it as global.
{
    If !wb    ;If wb is not a valid pointer then quit
        Return False
    Loop    ;Once it starts loading wait until completes
        Sleep,100
    Until (!wb.busy)
    Loop    ;optional check to wait for the page to completely load
        Sleep,100
    Until (wb.Document.Readystate = "Complete")
Return True
}

IEGet(Name="") {
	;Retrieve pointer to existing IE window/tab
	IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
		Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs" : RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer.*$" )
	For Pwb in ComObjCreate( "Shell.Application" ).Windows
		If (InStr(Pwb.LocationURL, Name)) && InStr(Pwb.FullName, "iexplore.exe")
			Return Pwb
}