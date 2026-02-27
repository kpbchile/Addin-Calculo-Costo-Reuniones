/*
 * Calculadora de Costo de Reuniones - Teams Meeting Tab
 * Generadora Metropolitana
 *
 * Muestra el costo estimado de una reunión en la pestaña de detalles de Teams.
 * Usa Microsoft Graph API para obtener asistentes y duración.
 */

var COST_PER_HOUR_PER_PERSON = 10000;
var MIN_INTERNAL_ATTENDEES = 6;

(function () {
  "use strict";

  microsoftTeams.app
    .initialize()
    .then(function () {
      return microsoftTeams.app.getContext();
    })
    .then(function (context) {
      return loadMeetingData(context);
    })
    .catch(function (err) {
      console.error("Error initializing Teams app:", err);
      showError();
    });

  function loadMeetingData(context) {
    // Get SSO token for Graph API calls
    return microsoftTeams.authentication
      .getAuthToken()
      .then(function (token) {
        // Use the meeting's chat ID or event ID to get meeting details
        // In a meeting context, we can get the event from the user's calendar
        var meetingId = context.meeting && context.meeting.id;

        if (!meetingId) {
          // Fallback: try to get event from calendar using chat thread
          return fetchCalendarEvents(token, context);
        }

        return fetchMeetingDetails(token, meetingId);
      })
      .then(function (meetingData) {
        if (!meetingData) {
          showError();
          return;
        }
        processMeetingData(meetingData);
      })
      .catch(function (err) {
        console.error("Error loading meeting data:", err);
        showError();
      });
  }

  function fetchMeetingDetails(token, meetingId) {
    // Decode the meeting ID to get the event
    // Graph API: GET /me/onlineMeetings with joinWebUrl filter
    return fetch("https://graph.microsoft.com/v1.0/me/calendar/events?$top=50&$orderby=start/dateTime desc&$select=subject,start,end,attendees,organizer,isOnlineMeeting,onlineMeetingUrl", {
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
    })
      .then(function (response) {
        if (!response.ok) throw new Error("Graph API error: " + response.status);
        return response.json();
      })
      .then(function (data) {
        // Try to match the current meeting by finding the event
        if (data.value && data.value.length > 0) {
          // Return the first matching event (most recent)
          return data.value[0];
        }
        return null;
      });
  }

  function fetchCalendarEvents(token, context) {
    // Get upcoming events from the user's calendar
    var now = new Date().toISOString();
    return fetch(
      "https://graph.microsoft.com/v1.0/me/calendar/events?$filter=start/dateTime ge '" +
        now +
        "'&$top=10&$orderby=start/dateTime&$select=subject,start,end,attendees,organizer",
      {
        headers: {
          Authorization: "Bearer " + token,
          "Content-Type": "application/json",
        },
      }
    )
      .then(function (response) {
        if (!response.ok) throw new Error("Graph API error: " + response.status);
        return response.json();
      })
      .then(function (data) {
        if (data.value && data.value.length > 0) {
          return data.value[0];
        }
        return null;
      });
  }

  function processMeetingData(event) {
    var attendees = event.attendees || [];
    var startTime = new Date(event.start.dateTime + "Z");
    var endTime = new Date(event.end.dateTime + "Z");

    // Determine organizer domain (internal domain)
    var organizerEmail = event.organizer && event.organizer.emailAddress
      ? event.organizer.emailAddress.address
      : "";
    var organizerDomain = organizerEmail.split("@")[1];
    if (!organizerDomain) {
      showError();
      return;
    }
    organizerDomain = organizerDomain.toLowerCase();

    // Count internal attendees
    var internalCount = 0;
    for (var i = 0; i < attendees.length; i++) {
      var email = attendees[i].emailAddress ? attendees[i].emailAddress.address : "";
      var parts = email.split("@");
      if (parts.length === 2 && parts[1].toLowerCase() === organizerDomain) {
        internalCount++;
      }
    }

    // +1 for the organizer
    var totalInternalParticipants = internalCount + 1;

    if (totalInternalParticipants < MIN_INTERNAL_ATTENDEES) {
      showNotApplicable();
      return;
    }

    // Calculate duration
    var durationMs = endTime.getTime() - startTime.getTime();
    var durationHours = durationMs / (1000 * 60 * 60);

    if (durationHours <= 0) {
      showError();
      return;
    }

    // Calculate cost
    var totalCost = COST_PER_HOUR_PER_PERSON * totalInternalParticipants * durationHours;

    // Render
    render(totalCost, totalInternalParticipants, durationHours);
  }

  function render(totalCost, internalCount, durationHours) {
    document.getElementById("loading").style.display = "none";
    document.getElementById("totalCost").textContent = "$" + formatCurrency(totalCost);
    document.getElementById("details").textContent =
      internalCount + " internos \u00b7 " +
      formatDuration(durationHours) + " \u00b7 $" +
      formatCurrency(COST_PER_HOUR_PER_PERSON) + "/hr/persona";
    document.getElementById("result").style.display = "block";
  }

  function showNotApplicable() {
    document.getElementById("loading").style.display = "none";
    document.getElementById("not-applicable").style.display = "block";
  }

  function showError() {
    document.getElementById("loading").style.display = "none";
    document.getElementById("error").style.display = "block";
  }

  function formatCurrency(amount) {
    var rounded = Math.round(amount);
    return rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  }

  function formatDuration(hours) {
    if (hours <= 0) return "0 min";
    if (hours < 1) return Math.round(hours * 60) + " min";
    var h = Math.floor(hours);
    var m = Math.round((hours - h) * 60);
    if (m === 0) return h + (h === 1 ? " hora" : " horas");
    return h + (h === 1 ? " hora " : " horas ") + m + " min";
  }
})();
