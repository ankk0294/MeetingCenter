var meeting_detail;
var badge_text = "MN";
var badge_show_timer = 0.1;
var badge_remove_timer = 0.1;

const showBadge = "showBadge";
const clearBadge = "clearBadge";

chrome.runtime.onStartup.addListener(
    () => {
        // TODO: Load meeting detail and set it to `meeting_detail`.
        startBadgeTimer();
    }
);

chrome.runtime.onMessage.addListener(
    function(request, sender, senderResponse) {
        meeting_detail = request;
        startBadgeTimer();
    }
);

function setBadgeAndTimers() {
    //TODO: Set `badge_text`, `badge_show_timer` and `badge_remove_timer` based on meeting_detail.
}

function cleanAlarms() {
    // Delete current alarms.
    chrome.alarms.clear(showBadge);
    chrome.alarms.clear(clearBadge);
}

function startBadgeTimer() {
    setBadgeAndTimers();
    cleanAlarms();
    chrome.alarms.create(showBadge, {delayInMinutes: badge_show_timer});
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
        }
    }
);