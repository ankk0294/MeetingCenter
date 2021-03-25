const TOKEN = "";
const URL = "https://graph.microsoft.com/v1.0/me/calendarview?startdatetime=2021-03-25T11:27:29.430Z&enddatetime=2021-03-30T11:27:29.430Z&orderby=start/dateTime"

fetch_data();

async function fetch_data() {
    var bearer = 'Bearer ' + TOKEN;
    var prefer = 'outlook.timezone="India Standard Time"';
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

function display_meeting_data(meeting_data) {
    if (!meeting_data.value) {
        return;
    }
    if (meeting_data.value.length == 0) {
        var no_meeting_div = document.createElement("div");
        var name_text = document.createTextNode("No meetings");
        no_meeting_div.appendChild(name_text);
        no_meeting_div.setAttribute("id", "no_meeting");
        document.body.appendChild(no_meeting_div);
        return;
    }
    // Create table.
    var table_element = document.createElement("table");
    table_element.setAttribute("id", "meeting_table");
    document.body.appendChild(table_element);

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
        if (!meeting_detail.isOnlineMeeting) {
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
        let date_format = new Date(meeting_detail.start.dateTime);
        var date_string = date_format.toLocaleDateString();
        var time_string = date_format.toLocaleTimeString();
        var trimmed_time_string = time_string.substring(0, time_string.length-6) + " " + time_string.substring(time_string.length-2, time_string.length);
        var start_time = document.createTextNode(date_string + "\n" + trimmed_time_string);
        time.appendChild(start_time);
        document.getElementById("row" + i).appendChild(time);
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

function start_timer(meeting_detail) {
    console.log(meeting_detail);
    chrome.runtime.sendMessage(JSON.stringify(meeting_detail));
}