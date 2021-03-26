var meeting_detail;
var badge_text_static = "Upcoming Meeting: ";
var badge_show_timer = 0.1;
var badge_remove_timer = 5;
var meeting_url = "";

const showBadge = "showBadge";
const clearBadge = "clearBadge";
const autoJoin = "autojoin";


chrome.runtime.onStartup.addListener(
    () => {
        // TODO: Load meeting detail and set it to `meeting_detail`.
    }
);

chrome.runtime.onMessage.addListener(
    function(request, sender, senderResponse) {
        meeting_detail = JSON.parse(request);
        startBadgeTimer(meeting_detail);
    }
);

function setBadgeAndTimers(meeting_detail) {
    badge_text = "";
    badge_text = badge_text_static+" "+meeting_detail.subject;
    var meeting_time = new Date(meeting_detail.start.dateTime)
    var current_time = new Date();
    if (meeting_time < current_time) {
       return;
    }
    badge_show_timer = new Date(meeting_time.getTime() - 5 * 60 * 1000).getTime();
    meeting_url = meeting_detail.onlineMeeting.joinUrl;
    chrome.alarms.create(autoJoin, {when: meeting_time.getTime()});
}

function cleanAlarms() {
    // Delete current alarms.
    chrome.alarms.clear(showBadge);
    chrome.alarms.clear(clearBadge);
}

function startBadgeTimer(meeting_detail) {
    setBadgeAndTimers(meeting_detail);
    cleanAlarms();
    chrome.alarms.create(showBadge, {when: badge_show_timer});
}

function setBadgeText(badge_text, badge_remove_timer) {
    chrome.browserAction.setBadgeText(
        {text: badge_text},
        () => {
            chrome.alarms.create(clearBadge, {delayInMinutes: badge_remove_timer});
        }
    );
}

function clearBadgeText() {
    chrome.browserAction.setBadgeText({text: ""});
}

chrome.alarms.onAlarm.addListener(
    (alarm) => {
        if (alarm.name == clearBadge) {
            clearBadgeText();
        } else if (alarm.name == showBadge) {
            setBadgeText(badge_text, badge_remove_timer);
        } else if (alarm.name == autoJoin) {
            chrome.tabs.create({
                url: meeting_url
           });
           clearBadgeText();
        }
    }
);