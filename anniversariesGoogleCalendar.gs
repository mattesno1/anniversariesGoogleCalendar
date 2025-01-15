/**
 * @file Anniversary Events Sync Script
 * 
 * @author github.com/JohnDOEbug
 * @version 1.0.0
 * 
 * @description
 * This script synchronizes birthdays and special events from Google
 * Contacts (contacts.google.com) with a selected Google Calendar
 * (calendar.google.com). It retrieves all relevant contact information
 * (name, events, resource ID) and checks it against existing calendar
 * entries. Entries are added or removed as needed to keep the calendar
 * up-to-date. Each calendar entry is set as a series with yearly events,
 * starting in the current year. The resource ID of each contact is stored
 * in the event description to ensure a unique link between the contact
 * and the event. Optionally, the user can receive email notifications for
 * the creation and deletion of events.
 * 
 * Usage:
 * - Create a new project for a Google Apps Script (script.google.com) 
 *   with your Google Account and copy this script.
 * - Set your Google Calendar ID and recipient email address in the 
 *   `anniversaryEvents` function.
 * - Add Google People API v1 to Services.
 * - Run the script once manually and accept the requested permissions.
 * - Add a Trigger for the function to run: anniversaryEvents. 
 *   (Suggestion - Event Source: Time-driven, Time based trigger: Day timer, 
 *   Time of Day: 8am to 9am)
 * 
 * Functions:
 * - anniversaryEvents: Main function to orchestrate the syncing process.
 * - getAllContactsEvents: Retrieves birthdays and other events from 
 *   Google Contacts.
 * - getAllCalendarEvents: Retrieves existing events from the specified 
 *   Google Calendar.
 * - compareAndSyncEvents: Compares contact and calendar events to determine 
 *   syncing actions.
 * - createCalendarEvent: Creates a recurring calendar event in Google 
 *   Calendar for the specified contact event.
 * - deleteCalendarEvent: Deletes a calendar event from Google Calendar.
 */

/**
 * Main function to synchronize anniversary events (birthdays and special 
 * events) between Google Contacts and Google Calendar. It retrieves all 
 * relevant events from Google Contacts, compares them with existing events 
 * in the specified Google Calendar, and updates the calendar by creating or 
 * deleting events as needed. At the end, a summary report of created and 
 * deleted events is emailed to the specified recipient, if any changes 
 * occurred and a recipient email is provided.
 *
 * @param {string} calendarId - The ID of the Google Calendar where events 
 *   are synchronized.
 * @param {string} recipientEmail - Optional email address to send a summary 
 *   report of created and deleted events. If empty, no email is sent.
 */
function anniversaryEvents() {
  var calendarId = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com'; // Set your Google Calendar ID here
  var recipientEmail = ''; // Set your E-Mail address here (optional)

  var currentYear = new Date().getFullYear();

  var contactEvents = getAllContactsEvents();
  var calendarEvents = getAllCalendarEvents(calendarId, currentYear);

  var syncActions = compareAndSyncEvents(contactEvents, calendarEvents);
  // Debug:
  //Logger.log("Sync Actions: %s", JSON.stringify(syncActions, null, 2));

  var createdEvents = [];
  var deletedEvents = [];

  syncActions.forEach(action => {
    if (action.eventId) {
      var deletedEventTitle = deleteCalendarEvent(calendarId, action.eventId);
      if (deletedEventTitle) {
        deletedEvents.push(deletedEventTitle);
      }
    } else {
      var createdEventTitle = createCalendarEvent(calendarId, action);
      if (createdEventTitle) {
        createdEvents.push(createdEventTitle);
      }
    }
  });

  if ((createdEvents.length > 0 || deletedEvents.length > 0) && recipientEmail) {
    createdEvents.sort();
    deletedEvents.sort();
    
    var reportMessage = 'Created Appointments:\n' + createdEvents.join('\n') + 
                        '\n\nDeleted Appointments:\n' + deletedEvents.join('\n');
    MailApp.sendEmail(recipientEmail, 'Anniversaries Google Calendar Report', reportMessage);
  }
}

/**
 * Retrieves all anniversary-related events (birthdays and other 
 * significant dates) from Google Contacts.
 * 
 * @returns {Array<Object>} Array of contact event objects with title, date, 
 *   and contactId properties.
 * @throws An error if no contacts retreived 
 */
function getAllContactsEvents() {
  var contactsEvents = [];
  var pageToken;
  var pageSize = 100;

  do {
    try {
      var response = People.People.Connections.list('people/me', {
        personFields: "names,birthdays,events",
        pageSize: pageSize,
        pageToken: pageToken
      });

      var connections = response.connections;
      if (connections) {
        connections.forEach(function(contact) {
          var name = contact.names && contact.names[0] ? contact.names[0].displayName : "Unknown Name";
          var currentYear = new Date().getFullYear() // Just in case no year is defined in contact

          if (contact.birthdays) {
            contact.birthdays.forEach(function(birthday) {
              if (birthday.date) {
                var year = birthday.date.year || currentYear;
                var eventDate = new Date(year, birthday.date.month - 1, birthday.date.day)
                contactsEvents.push({
                  title: name + " hat Geburtstag",
                  date: eventDate,
                  contactId: contact.resourceName,
                });
              }
            });
          }

          if (contact.events) {
            contact.events.forEach(function(event) {
              if (event.date && event.formattedType) {
                var year = event.date.year || currentYear;
                var eventDate = new Date(year, event.date.month - 1, event.date.day)
                contactsEvents.push({
                  title: event.formattedType + ": " + name,
                  date: eventDate,
                  contactId: contact.resourceName,
                });
              }
            });
          }
        });
      }

      pageToken = response.nextPageToken;
    } catch (error) {
      Logger.log("Error retrieving contacts: " + error);
      break;
    }
  } while (pageToken);

  return contactsEvents;
}

/**
 * Retrieves all events from a Google Calendar for a specific year, 
 * filtered by events that match the contact resource name pattern.
 * 
 * @param {string} calendarId - The ID of the Google Calendar to retrieve 
 *   events from.
 * @param {number} year - The year for which events should be retrieved.
 * @returns {Array<Object>} Array of calendar event objects with title, 
 *   date, contactId, and eventId properties.
 */
function getAllCalendarEvents(calendarId, year) {
  var calendar = CalendarApp.getCalendarById(calendarId);

  var startDate = new Date(year, 0, 1);
  var endDate = new Date(year + 1, 0, 1);

  const resourceNamePattern = /people\/c\d+/;

  return calendar.getEvents(startDate, endDate)
    .filter(event => {
      const description = event.getDescription();
      return description && resourceNamePattern.test(description);
    })
    .map(event => {
      const description = event.getDescription();
      const matchedResourceName = description.match(resourceNamePattern)[0];

      return {
        title: event.getTitle(),
        date: event.getStartTime(),
        contactId: matchedResourceName,
        eventId: event.getId()
      };
    });
}

/**
 * Compares contact and calendar events to identify new events to add 
 * and obsolete events to delete.
 * 
 * @param {Array<Object>} contactEvents - Array of contact events retrieved 
 *   from Google Contacts.
 * @param {Array<Object>} calendarEvents - Array of calendar events 
 *   retrieved from Google Calendar.
 * @returns {Array<Object>} Array of sync action objects specifying which 
 *   events to add or delete.
 */
function compareAndSyncEvents(contactEvents, calendarEvents) {
  const syncActions = [];

  function isSameDayAndMonth(date1, date2) {
    const d1 = new Date(date1);
    const d2 = new Date(date2);
    return d1.getDate() === d2.getDate() && d1.getMonth() === d2.getMonth();
  }

  contactEvents.forEach(contactEvent => {
    const matchingCalendarEvent = calendarEvents.find(calendarEvent =>
      calendarEvent.title === contactEvent.title &&
      isSameDayAndMonth(calendarEvent.date, contactEvent.date) &&
      calendarEvent.contactId === contactEvent.contactId
    );

    if (!matchingCalendarEvent) {
      syncActions.push(contactEvent);
    }
  });

  calendarEvents.forEach(calendarEvent => {
    const matchingContactEvent = contactEvents.find(contactEvent =>
      contactEvent.title === calendarEvent.title &&
      isSameDayAndMonth(contactEvent.date, calendarEvent.date) &&
      contactEvent.contactId === calendarEvent.contactId
    );

    if (!matchingContactEvent) {
      syncActions.push(calendarEvent);
    }
  });

  return syncActions;
}

/**
 * Creates a recurring all-day event in Google Calendar for the given 
 * contact event.
 * 
 * @param {string} calendarId - The ID of the Google Calendar to add the 
 *   event to.
 * @param {Object} event - The event object containing title, date, and 
 *   contactId.
 * @returns {string} The title of the created event.
 */
function createCalendarEvent(calendarId, event) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var title = event.title;
  var description = event.contactId;

  var currentYear = new Date().getFullYear();
  var startDate = new Date(currentYear, event.date.getMonth(), event.date.getDate());

  calendar.createAllDayEventSeries(
    title,
    startDate, 
    CalendarApp.newRecurrence().addYearlyRule(),
    { description: description }
  );

  return title;
}

/**
 * Deletes an event from Google Calendar.
 * 
 * @param {string} calendarId - The ID of the Google Calendar to delete the 
 *   event from.
 * @param {string} eventId - The ID of the event to be deleted.
 * @returns {string} The title of the deleted event.
 */
function deleteCalendarEvent(calendarId, eventId) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var eventToDelete = calendar.getEventById(eventId);

  if (eventToDelete) {
    var title = eventToDelete.getTitle();
    eventToDelete.deleteEvent();
    return title;
  } else {
    Logger.log('Event not found: ' + eventId);
    return null;
  }
}
