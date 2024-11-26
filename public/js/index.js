// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// get value from html here

$("#embeddedButton").click(() => {
    const type = $("#type").val();
    const sourceId = $("#sourceId").val();
    const workspaceId = $("#workspaceId").val();

    const embeddedHistoryList = JSON.parse(localStorage.getItem("embeddedHistoryList") || "[]");
    console.log('embeddedHistoryList 1', embeddedHistoryList);

    embeddedHistoryList.push(JSON.stringify({
        timestamp: Date.now(),
        type,
        workspaceId,
        sourceId,
    }))

    localStorage.setItem("embeddedHistoryList", JSON.stringify(embeddedHistoryList));

    printLogTable(embeddedHistoryList)

    callGenEmbedToken({
        type,
        sourceId,
        workspaceId,
    })

    $("#type").val();
    $("#sourceId").val("");
    $("#workspaceId").val("");
});

const printLogTable = (logList = []) => {

    $("#logTableBody").empty()

    logList.forEach((jsonLog) => {
        const log = JSON.parse(jsonLog);

        $("#logTableBody").append(`
            <tr>
                <td> ${log.type} </td>
                <td> ${log.workspaceId} </td>
                <td> ${log.sourceId} </td>
                <td> ${new Date(log.timestamp).toLocaleString("en-gb").split(", ").join(" ")} </td>
                <td>
                    <button 
                        class="btn btn-info embed-log-btn"
                        data-log='${JSON.stringify(log)}'
                    > 
                        embed
                        <i class="fa-solid fa-moon"></i>
                    </button>
                    <button 
                        class="btn btn-danger delete-log-btn" 
                        data-timestamp="${log.timestamp}"
                    > 
                        x 
                    </button>
                </td>
            </tr>        
        `)
    })
}

$(document).on("click", ".embed-log-btn", function() {
    const log = $(this).data("log")

    callGenEmbedToken({
        type: log.type,
        sourceId: log.sourceId,
        workspaceId: log.workspaceId,
    })
})

$(document).on("click", ".delete-log-btn", function() {
    const logList = JSON.parse(localStorage.getItem("embeddedHistoryList") || "[]");
    const timestamp = $(this).data("timestamp");

    const updateLog = logList.filter(stringLog => {
        const log = JSON.parse(stringLog);
        return log.timestamp !== timestamp
    })

    localStorage.setItem("embeddedHistoryList", JSON.stringify(updateLog))
    printLogTable(updateLog)
})

printLogTable(JSON.parse(localStorage.getItem("embeddedHistoryList") || "[]"))

let models = window["powerbi-client"].models;
let reportContainer = $("#report-container").get(0);

// Initialize iframe for embedding report
powerbi.bootstrap(reportContainer, { type: "dashboard" });

const callGenEmbedToken = (form) => {
    // AJAX request to get the report details from the API and pass it to the UI
    console.log("form", form)
    $.ajax({
        type: "POST",
        url: "/getEmbedToken",
        dataType: "json",
        data: form,
        success: function (embedData) {
            // Create a config object with type of the object, Embed details and Token Type
            let reportLoadConfig = {
                type: "dashboard",
                tokenType: models.TokenType.Embed,
                accessToken: embedData.accessToken,

                // Use other embed report config based on the requirement. We have used the first one for demo purpose
                embedUrl: embedData.embedUrl[0].embedUrl,

                // Enable this setting to remove gray shoulders from embedded report
                // settings: {
                //     background: models.BackgroundType.Transparent
                // }
            };

            // Use the token expiry to regenerate Embed token for seamless end user experience
            // Refer https://aka.ms/RefreshEmbedToken
            tokenExpiry = embedData.expiry;

            // Embed Power BI report when Access token and Embed URL are available
            let report = powerbi.embed(reportContainer, reportLoadConfig);

            // Clear any other loaded handler events
            report.off("loaded");

            // Triggers when a report schema is successfully loaded
            report.on("loaded", function () {
                console.log("Report load successful");
            });

            // Clear any other rendered handler events
            report.off("rendered");

            // Triggers when a report is successfully embedded in UI
            report.on("rendered", function () {
                console.log("Report render successful");
            });

            // Clear any other error handler events
            report.off("error");

            // Handle embed errors
            report.on("error", function (event) {
                let errorMsg = event.detail;
                console.error(errorMsg);
                return;
            });
        },

        error: function (err) {
            // Show error container
            let errorContainer = $(".error-container");
            $(".embed-container").hide();
            errorContainer.show();

            // Get the error message from err object
            let errMsg = JSON.parse(err.responseText)["error"];

            // Split the message with \r\n delimiter to get the errors from the error message
            let errorLines = errMsg.split("\r\n");

            // Create error header
            let errHeader = document.createElement("p");
            let strong = document.createElement("strong");
            let node = document.createTextNode("Error Details:");

            // Get the error container
            let errContainer = errorContainer.get(0);

            // Add the error header in the container
            strong.appendChild(node);
            errHeader.appendChild(strong);
            errContainer.appendChild(errHeader);

            // Create <p> as per the length of the array and append them to the container
            errorLines.forEach((element) => {
                let errorContent = document.createElement("p");
                let node = document.createTextNode(element);
                errorContent.appendChild(node);
                errContainer.appendChild(errorContent);
            });
        },
    });
};
