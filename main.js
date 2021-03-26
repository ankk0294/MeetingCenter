const TOKEN = "";
const NEW_MEETING_URL = "https://graph.microsoft.com/v1.0/me/events"

fetch_data()

async function fetch_data() {
    var bearer = 'Bearer ' + TOKEN;
    var prefer = 'outlook.timezone="India Standard Time"';
    var start_time = new Date();
    var end_time = new Date(start_time.getTime() + 24 * 60 * 60 * 1000);
    const URL = "https://graph.microsoft.com/v1.0/me/calendarview?startdatetime="+ start_time.toISOString() +"&enddatetime=" + end_time.toISOString() +"&orderby=start/dateTime"
    let data = await fetch(URL, {
        method: 'GET',
        headers: {
            'Authorization': bearer,
            'Prefer': prefer
        }
    });
    let response_data = await data.json();
    display_meeting_data(response_data);
}

function createMeeting() {
    var request = new XMLHttpRequest();
    var bearer = 'Bearer ' + TOKEN;
    var prefer = 'outlook.timezone="India Standard Time"';
    request.open("POST", NEW_MEETING_URL, true);
    request.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
    request.setRequestHeader("Authorization", bearer);
    request.setRequestHeader("Prefer", prefer);
    var element = document.getElementById('subject');
    var subject = element.value;
    element = document.getElementById('content');
    var content = element.value;
    element = document.getElementById('receipent');
    var receipent = element.value;
    element = document.getElementById('start_time');
    var start_time = element.value;
    element = document.getElementById('end_time');
    var end_time = element.value;
    request.send(JSON.stringify({
        "subject": subject,
        "body": {
        "contentType": "HTML",
          "content": content
            },
            "start": {
                "dateTime": start_time,
                "timeZone": "India Standard Time"
            },
            "end": {
                "dateTime": end_time,
                "timeZone": "India Standard Time"
            },
            "location":{
                "displayName":"Teams"
            },
            "attendees": [
              {
                "emailAddress": {
                  "address":receipent
                },
       "type": "required"
           }
        ],
        "allowNewTimeProposals": true,
        "transactionId":"7E163156-7762-4BEB-A1C6-729EA81755A7"
    }));
    fetch_data();
}

function addMeeting() {
    var root_element = document.getElementById("root");
    if (root_element != undefined) {
        root_element.innerHTML = "";
    }

    var start_time_div = document.createElement("div");
    var start_time_text = document.createTextNode("Start Time");
    start_time_div.appendChild(start_time_text);
    var start_time = document.createElement("Input");
    start_time.setAttribute("id", "start_time");
    start_time.setAttribute("type", "datetime-local");
    start_time_div.appendChild(start_time);

    var end_time_div = document.createElement("div");
    var end_time_text = document.createTextNode("End Time");
    end_time_div.appendChild(end_time_text);
    var end_time = document.createElement("Input");
    end_time.setAttribute("id", "end_time");
    end_time.setAttribute("type", "datetime-local");
    end_time_div.appendChild(end_time);

    var subject_div = document.createElement("div");
    var subject_text = document.createTextNode("Subject");
    subject_div.appendChild(subject_text);
    var subject = document.createElement("Input");
    subject.setAttribute("id", "subject");
    subject.setAttribute("type", "text");
    subject_div.appendChild(subject);

    var text_area_div = document.createElement("div");
    var text_area_text = document.createTextNode("Message Content");
    text_area_div.appendChild(text_area_text);
    var text_area = document.createElement("Input");
    text_area.setAttribute("id", "content");
    text_area.setAttribute("type", "text");
    text_area_div.appendChild(text_area);

    var receipent_div = document.createElement("div");
    var receipent_text = document.createTextNode("Receipient");
    receipent_div.appendChild(receipent_text);
    var receipent = document.createElement("Input");
    receipent.setAttribute("id", "receipent");
    receipent.setAttribute("type", "email");
    receipent_div.appendChild(receipent);

    var send = document.createElement("Input");
    send.setAttribute("id", "send");
    send.setAttribute("type", "button");
    send.setAttribute("value", "Send");

    var cancel = document.createElement("Input");
    cancel.setAttribute("id", "cancel");
    cancel.setAttribute("type", "button");
    cancel.setAttribute("value", "Cancel");

    root_element.appendChild(start_time_div);
    root_element.appendChild(end_time_div);
    root_element.appendChild(subject_div);
    root_element.appendChild(text_area_div);
    root_element.appendChild(receipent_div);
    root_element.appendChild(send);
    root_element.appendChild(cancel);

    var sendButton = document.getElementById('send');
    sendButton.addEventListener('click', function() {
        createMeeting();
    }, false);
    var cancelButton = document.getElementById('cancel');
    cancelButton.addEventListener('click', function() {
        fetch_data();
    }, false);
    // fetch_data();
}

function display_meeting_data(meeting_data) {
    if (!meeting_data.value) {
        return;
    }
    var root_element = document.getElementById("root");
    if (root_element != undefined) {
        root_element.innerHTML = "";
    }
    
    var add_meeting = document.createElement("Input");
    add_meeting.setAttribute("id", "addMeeting");
    add_meeting.setAttribute("type", "button");
    add_meeting.setAttribute("value", "Add Meeting");
    root_element.appendChild(add_meeting);

    var auto_join_div = document.createElement("div");
    auto_join_div.setAttribute("id", "autoJoin");
    var autojoin_text = document.createTextNode("Auto Join");
    auto_join_div.appendChild(autojoin_text);
    var auto_join_check_box = document.createElement("Input");
    auto_join_check_box.setAttribute("id", "autoJoinCheck");
    auto_join_check_box.setAttribute("type", "checkbox");
    auto_join_check_box.setAttribute("checked","true");
    auto_join_div.appendChild(auto_join_check_box);
    root_element.appendChild(auto_join_div);

    var checkPageButton = document.getElementById('addMeeting');
    checkPageButton.addEventListener('click', function() {
        addMeeting();
    }, false);

    if (meeting_data.value.length == 0) {
        var no_meeting_div = document.createElement("div");
        no_meeting_div.setAttribute("id", "NoMeeting");
        var name_text = document.createTextNode("No meetings");
        no_meeting_div.appendChild(name_text);
        no_meeting_div.setAttribute("id", "no_meeting");
        root_element.appendChild(no_meeting_div);
        return;
    }
    // Create table.
    var table_element = document.createElement("table");
    table_element.setAttribute("id", "meeting_table");
    root_element.appendChild(table_element);

    // Create table headings.
    var heading = document.createElement("tr");
    heading.setAttribute("id", "table_heading");
    document.getElementById("meeting_table").appendChild(heading);
    // Create columns in heading.
    var meeting_name = document.createElement("th");
    var name_text = document.createTextNode("Meeting Name");
    meeting_name.appendChild(name_text);
    document.getElementById("table_heading").appendChild(meeting_name);
    var meeting_time = document.createElement("th");
    var time_text = document.createTextNode("Meeting time");
    meeting_time.appendChild(time_text);
    document.getElementById("table_heading").appendChild(meeting_time);
    var meeting_url = document.createElement("th");
    var url_text = document.createTextNode("Meeting URL");
    meeting_url.appendChild(url_text);
    document.getElementById("table_heading").appendChild(meeting_url);

    // Create rows in the table.
    var sent_first_meeting = false;
    const row_count = 5;
    var current_count = 0;
    for (var i = 0; i < meeting_data.value.length; i++) {
        let meeting_detail = meeting_data.value[i];
        let date_format = new Date(meeting_detail.start.dateTime);
        var current_time = new Date();
        if (date_format < current_time) {
            continue;
        }
        current_count += 1;
        if (current_count > row_count) {
            break;
        }
        if (!sent_first_meeting) {
            start_timer(meeting_detail);
            sent_first_meeting = true;
        }
        var row = document.createElement("tr");
        row.setAttribute("id", "row" + i);
        document.getElementById("meeting_table").appendChild(row);
        // Update values for each column.
        var name = document.createElement("td");
        var subject = document.createTextNode(meeting_detail.subject);
        name.appendChild(subject);
        document.getElementById("row" + i).appendChild(name);
        var time = document.createElement("td");
        var date_string = date_format.toLocaleDateString();
        var time_string = date_format.toLocaleTimeString();
        var trimmed_time_string = time_string.substring(0, time_string.length-6) + " " + time_string.substring(time_string.length-2, time_string.length);
        var start_time = document.createTextNode(date_string + "\n" + trimmed_time_string);
        time.appendChild(start_time);
        document.getElementById("row" + i).appendChild(time);
        if (meeting_detail.isOnlineMeeting) {
            var url = document.createElement("td");
            var link_element = document.createElement("a");
            var join_url = document.createTextNode("Join URL");
            link_element.appendChild(join_url);
            link_element.title = "Join URL";
            link_element.setAttribute("class", "button");
            link_element.href = meeting_detail.onlineMeeting.joinUrl;
            url.appendChild(link_element);
            document.getElementById("row" + i).appendChild(url);
        }
    }

    if (current_count == 0) {
        var no_meeting_div = document.createElement("div");
        no_meeting_div.setAttribute("id", "NoMeeting");
        var name_text = document.createTextNode("No meetings");
        no_meeting_div.appendChild(name_text);
        no_meeting_div.setAttribute("id", "no_meeting");
        root_element.appendChild(no_meeting_div);
        return;
    }
}

function start_timer(meeting_detail) {
    chrome.runtime.sendMessage(JSON.stringify(meeting_detail));
}
