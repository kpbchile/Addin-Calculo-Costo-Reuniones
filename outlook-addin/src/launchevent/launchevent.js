/*
 * Calculadora de Costo de Reuniones - Generadora Metropolitana
 *
 * - OnAppointmentSend: muestra Smart Alert con costo antes de enviar
 * - calculateCostFunction: boton manual para mostrar InfoBar
 */

var COST_PER_HOUR_PER_PERSON = 10000;
var MIN_INTERNAL_ATTENDEES = 6;
var NOTIFICATION_KEY = "costReunion";

// --- Smart Alert: se ejecuta al presionar Enviar ---

function onAppointmentSendHandler(event) {
  gatherMeetingData(function (data) {
    if (!data) {
      event.completed({ allowEvent: true });
      return;
    }

    var costData = computeCost(data);

    if (costData.totalInternalParticipants < MIN_INTERNAL_ATTENDEES) {
      event.completed({ allowEvent: true });
      return;
    }

    var formattedCost = formatCurrency(costData.totalCost);
    var durationDisplay = formatDuration(costData.durationHours);

    var message =
      "Costo estimado de esta reunion: $" + formattedCost +
      "\n\nParticipantes internos: " + costData.totalInternalParticipants +
      "\nDuracion: " + durationDisplay +
      "\nTarifa por persona por hora: $" + formatCurrency(COST_PER_HOUR_PER_PERSON) +
      "\n\nDesea enviar la invitacion de todas formas?";

    event.completed({
      allowEvent: false,
      errorMessage: message,
    });
  });
}

// --- Boton ribbon: muestra InfoBar con costo ---

function calculateCostFunction(event) {
  gatherMeetingData(function (data) {
    if (!data) {
      showNotification("No se pudo obtener la informacion de la reunion.", event);
      return;
    }

    var costData = computeCost(data);

    if (costData.totalInternalParticipants < MIN_INTERNAL_ATTENDEES) {
      showNotification(
        "Esta reunion tiene " + costData.totalInternalParticipants +
        " participantes internos. El calculo se activa con 6 o mas.", event);
      return;
    }

    var formattedCost = formatCurrency(costData.totalCost);
    var durationDisplay = formatDuration(costData.durationHours);
    var infoMessage =
      "Costo reunion: $" + formattedCost +
      " (" + costData.totalInternalParticipants + " internos, " +
      durationDisplay + ", $" + formatCurrency(COST_PER_HOUR_PER_PERSON) + "/hr/persona)";

    showNotification(infoMessage, event);
  });
}

function showNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    NOTIFICATION_KEY,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "Icon.16x16",
      persistent: true,
    },
    function () {
      event.completed();
    }
  );
}

// --- Data gathering ---

function gatherMeetingData(callback) {
  var userEmail = Office.context.mailbox.userProfile.emailAddress;
  var userDomain = userEmail.split("@")[1].toLowerCase();

  var result = {
    userDomain: userDomain,
    requiredAttendees: [],
    optionalAttendees: [],
    startTime: null,
    endTime: null,
  };

  var pending = 4;

  function checkDone() {
    pending--;
    if (pending === 0) {
      if (!result.startTime || !result.endTime) {
        callback(null);
      } else {
        callback(result);
      }
    }
  }

  Office.context.mailbox.item.requiredAttendees.getAsync(function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      result.requiredAttendees = res.value;
    }
    checkDone();
  });

  Office.context.mailbox.item.optionalAttendees.getAsync(function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      result.optionalAttendees = res.value;
    }
    checkDone();
  });

  Office.context.mailbox.item.start.getAsync(function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      result.startTime = res.value;
    }
    checkDone();
  });

  Office.context.mailbox.item.end.getAsync(function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      result.endTime = res.value;
    }
    checkDone();
  });
}

function computeCost(data) {
  var allAttendees = data.requiredAttendees.concat(data.optionalAttendees);
  var internalCount = 0;

  for (var i = 0; i < allAttendees.length; i++) {
    var email = allAttendees[i].emailAddress || "";
    var parts = email.split("@");
    if (parts.length === 2 && parts[1].toLowerCase() === data.userDomain) {
      internalCount++;
    }
  }

  var totalInternalParticipants = internalCount + 1;

  var durationMs = data.endTime.getTime() - data.startTime.getTime();
  var durationHours = durationMs / (1000 * 60 * 60);

  if (durationHours <= 0) {
    durationHours = 0;
  }

  var totalCost = COST_PER_HOUR_PER_PERSON * totalInternalParticipants * durationHours;

  return {
    totalInternalParticipants: totalInternalParticipants,
    durationHours: durationHours,
    totalCost: totalCost,
  };
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

// --- Auto-launch: se ejecuta al abrir nueva reunion ---

function onNewAppointmentOrganizer(event) {
  // Mostrar InfoBar inicial informando que el calculo esta activo
  showNotification(
    "Calculadora de costo activa. Agregue asistentes y presione Enviar para ver el costo.",
    event
  );
}

// --- Register handlers ---

Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate("calculateCostFunction", calculateCostFunction);
