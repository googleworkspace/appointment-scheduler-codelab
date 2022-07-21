/**
 * Copyright 2022 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

'use strict';

// Import the Dialogflow module from Google client libraries.
const functions = require('firebase-functions');
const {google} = require('googleapis');
const {WebhookClient, Payload} = require('dialogflow-fulfillment');

// Enter your calendar ID below and service account JSON below
 // Enter your calendar ID below and service account JSON below
 const calendarId = '<Add your calendar ID here>';
 const serviceAccount = {<Add your service account details here>}; // Starts with {"type": "service_account",...
 
const API_KEY = 'YOUR_API_KEY';
const MAP_IMAGE_URL = 'https://maps.googleapis.com/maps/api/staticmap?center=Googleplex&zoom=14&size=200x200&key=' + API_KEY;
const ICON_IMAGE_URL = 'https://fonts.gstatic.com/s/i/googlematerialicons/calendar_today/v5/black-48dp/1x/gm_calendar_today_black_48dp.png';
const CALENDAR_URL = 'YOUR_CALENDAR_URL';

// Set up Google Calendar Service account credentials
const serviceAccountAuth = new google.auth.JWT({
  email: serviceAccount.client_email,
  key: serviceAccount.private_key,
  scopes: 'https://www.googleapis.com/auth/calendar',
});

const calendar = google.calendar('v3');
process.env.DEBUG = 'dialogflow:*'; // enables lib debugging statements

const timeZone = 'America/Los_Angeles';
const timeZoneOffset = '-07:00';

// Set the DialogflowApp object to handle the HTTPS POST request.
exports.dialogflowFirebaseFulfillment = functions.https.onRequest(
    (request, response) => {
      const agent = new WebhookClient({request, response});
      console.log('Parameters', agent.parameters);
      const appointmentType = agent.parameters.AppointmentType;
      console.log(`fulfillment - appointment type: ${appointmentType}`);

      /**
       * Schedules the appointment.
       * @param {*} agent The Dialogflow agent.
       * @return {Promise} The event creation.
       */
      function makeAppointment(agent) {
        // Calculate appointment start and end datetimes (end = +1hr from start)
        const dateTimeStart = new Date(Date.parse(agent.parameters.date
            .split('T')[0] + 'T' + agent.parameters.time.split('T')[1].split('-')[0] + timeZoneOffset));
        const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
        const appointmentTimeString = dateTimeStart.toLocaleString(
            'en-US',
            {month: 'long', day: 'numeric', hour: 'numeric', timeZone: timeZone},
        );
        // Check the availability of the time, and make an appointment if there is time on the calendar
        return createCalendarEvent(dateTimeStart, dateTimeEnd, appointmentType).then(() => {
          console.log('Creating calendar event');
          agent.add(`Ok, let me see if we can fit you in. ${appointmentTimeString} is fine!.`);

          console.log(`dateTimeStart: ${dateTimeStart}`);
          console.log(`appointment time string: ${appointmentTimeString}`);
          const dateString = dateTimeStart.toLocaleString(
            'en-US',
              {month: 'long', day: 'numeric'},
          );
          const dateParts = appointmentTimeString.split(',');
          console.log(`dateString: ${dateString}`);
          console.log(`dateParts: ${dateParts}`);
          const json = getHangoutsCard(appointmentType, dateParts[0], dateParts[1]);
          const payload = new Payload(
              'hangouts',
              json,
              {rawPayload: true, sendAsMessage: true},
          );
          agent.add(payload);

        }).catch(() => {
          agent.add(`I'm sorry, there are no slots available for ${appointmentTimeString}.`);
        });
      }

      // Handle the Dialogflow intent named 'Schedule Appointment'.
      const intentMap = new Map();
      intentMap.set('Schedule Appointment', makeAppointment);
      agent.handleRequest(intentMap);
    });

/**
 * Creates calendar event in Google Calendar
 *
 * @param {Date} dateTimeStart The start time of the event.
 * @param {Date} dateTimeEnd The end time of the event.
 * @param {String} appointmentType The appointment activity.
 * @return {Promise} The event creation.
 */
function createCalendarEvent(dateTimeStart, dateTimeEnd, appointmentType) {
  return new Promise((resolve, reject) => {
    calendar.events.list({
      auth: serviceAccountAuth, // List events for time period
      calendarId: calendarId,
      timeMin: dateTimeStart.toISOString(),
      timeMax: dateTimeEnd.toISOString(),
    }, (err, calendarResponse) => {
      // Check if there is a event already on the Calendar
      if (err || calendarResponse.data.items.length > 0) {
        reject(err || new Error('Requested time conflicts with another appointment'));
      } else {
        // Create event for the requested time period
        calendar.events.insert({auth: serviceAccountAuth,
          calendarId: calendarId,
          resource: {summary: appointmentType +' Appointment', description: appointmentType,
            start: {dateTime: dateTimeStart},
            end: {dateTime: dateTimeEnd}},
        }, (err, event) => {
         err ? reject(err) : resolve(event);
        },
        );
      }
    });
  });
}

/**
 * Return a well formed test card
 *
 * @param {String} appointmentType The appointment activity.
 * @param {String} date The date of the event.
 * @param {String} time The time of the event.
 * @return {Object} JSON rich card
 */
function getHangoutsCard(appointmentType, date, time) {
  const cardHeader = {
    title: 'Appointment Confirmation',
    subtitle: appointmentType,
    imageUrl: ICON_IMAGE_URL,
    imageStyle: 'IMAGE',
  };

  const dateWidget = {
    keyValue: {
      content: 'Date',
      bottomLabel: date,
    },
  };

  const timeWidget = {
    keyValue: {
      content: 'Time',
      bottomLabel: time,
    },
  };

  const mapImageWidget = {
    'image': {
      'imageUrl': MAP_IMAGE_URL,
      'onClick': {
        'openLink': {
          'url': MAP_IMAGE_URL,
        },
      },
    },
  };

  const buttonWidget = {
    buttons: [
      {
        textButton: {
          text: 'View Appointment',
          onClick: {
            openLink: {
              url: CALENDAR_URL,
            },
          },
        },
      },
    ],
  };

  const infoSection = {widgets: [dateWidget, timeWidget, 
    buttonWidget]};

  return {
    'hangouts': {
      'name': 'Confirmation Card',
      'header': cardHeader,
      'sections': [infoSection],
    },
  };
}
