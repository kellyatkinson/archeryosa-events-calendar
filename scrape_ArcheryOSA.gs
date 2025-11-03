// This script scrapes ArcheryOSA upcoming event details and inserts them into a Google Calendar.
// It is suitable for daily use.
// Event updates are reflected when the script runs.
// Kelly Atkinson, 2025

const CONFIG = {
  CAL_ID: PropertiesService.getScriptProperties().getProperty('CALENDAR_ID'),
  BASE_URL: 'https://www.archeryosa.com'
};

function automateEventCreation() {
  const mainPageUrl = CONFIG.BASE_URL;
  const html = UrlFetchApp.fetch(mainPageUrl, ); // Fetch the main event listing page
  
  const events = parseEventData(html.getContentText()); // Parse the events from the table
  
  events.forEach(event => {
    createOrUpdateCalendarEvent(event); // Create Google Calendar events for each
  });
}

function fetchEventDetails(eventUrl) {
  const response = UrlFetchApp.fetch(eventUrl); // Fetch the event details page
  const html = response.getContentText();

  // Extract Start Date, End Date, and Host Club using regex
  const startDateMatch = html.match(/<th[^>]*>Start Date<\/th>\s*<td>([^<]+)<\/td>/);
  const endDateMatch = html.match(/<th[^>]*>End Date<\/th>\s*<td>([^<]+)<\/td>/);
  const hostClubMatch = html.match(/<th[^>]*>Host Club<\/th>\s*<td>([^<]+)<\/td>/);

  const startDate = startDateMatch ? startDateMatch[1].trim() : null;
  const endDate = endDateMatch ? endDateMatch[1].trim() : startDate; // Use start date if no end date found
  const hostClub = hostClubMatch ? hostClubMatch[1].trim() : 'Unknown Club';

  return { startDate, endDate, hostClub };
}

function parseEventData(html) {
  const events = [];
  
  // Regex to capture the event URL, name, date, region, and type from the main table
  const tableRowRegex = /<tr>\s*<th scope="row">\s*<a href="([^"]+)">\s*([^<]+)<\/a>\s*<\/th>\s*<td>([^<]+)<\/td>\s*<td>([^<]+)<\/td>\s*<td>([^<]+)<\/td>\s*<\/tr>/g;
  
  let match;
  while ((match = tableRowRegex.exec(html)) !== null) {
    const eventUrl = CONFIG.BASE_URL + match[1];  // Full event URL
    const eventName = match[2].trim();  // Event name
    const eventDate = match[3].trim();  // Event start date (e.g., "16 Nov")
    const eventRegion = match[4].trim();  // Event region
    const eventType = match[5].trim();  // Event type

    // Fetch additional details from the event page
    const details = fetchEventDetails(eventUrl);
    const startDate = details.startDate || eventDate;  // Use fetched start date if available
    const endDate = details.endDate || startDate;  // Use fetched end date if available
    const hostClub = details.hostClub || '';  // Use fetched host club if available
    if(hostClub == "Unknown Club") { const hostClub = ""};

    // Push the event's details into the events array
    events.push({
      url: eventUrl,
      name: eventName,
      startDate: startDate,
      endDate: endDate,
      region: eventRegion,
      type: eventType,
      hostClub: hostClub
    });
  }
  
  return events;
}

function createOrUpdateCalendarEvent(event) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CAL_ID);
  
  if (!calendar) {
    Logger.log('Calendar not found!');
    return;
  }

  const startDate = new Date(event.startDate);
  const endDate = new Date(event.endDate);
  
  // Set the end date to the next day at midnight to ensure it's an all-day event
  endDate.setDate(endDate.getDate() + 1); // This makes the event span the full end date (next day at 12:00 AM)

  // Make sure both start and end dates do not have times associated with them
  // startDate.setHours(0, 0, 0, 0); // Reset to midnight to ensure it's an all-day event
  // endDate.setHours(0, 0, 0, 0);   // Reset to midnight for the same reason

  // Search for existing events on the same day with the same event URL in the description
  const existingEvents = calendar.getEventsForDay(startDate)
    .filter(e => e.getDescription().includes(event.url));

  let eventTitle;

  // Modify event title based on the hostClub and event type
  if (event.hostClub === "Unknown Club" || event.type === "League") {
    eventTitle = `${event.name}`;  // No host club in the title if it's "Unknown Club" or "League"
  } else {
    // Remove "Archery Club" or "Archers" from the host club name if they exist
    let hostClubCleaned = event.hostClub
      .replace(/Archery Club/i, '') // Remove "Archery Club" (case insensitive)
      .replace(/Archers/i, '')       // Remove "Archers" (case insensitive)
      .trim();                      // Remove any extra spaces

    // If there are multiple words left in the hostClub, use the first two words, otherwise just use one word
    let hostClubWords = hostClubCleaned.split(' ').filter(word => word.length > 0); // Split and remove empty words

    if (hostClubWords.length > 1) {
      eventTitle = `${hostClubWords[0]} ${hostClubWords[1]}: ${event.name}`;  // Use the first two words
    } else if (hostClubWords.length === 1) {
      eventTitle = `${hostClubWords[0]}: ${event.name}`; // Use the single word
    } else {
      eventTitle = `${event.name}`;  // If no host club name remains, just use the event name
    }
  }

  if (existingEvents.length > 0) {
    // Update the existing event
    const existingEvent = existingEvents[0]; // Assuming there is only one event for the URL
    
    // Check if any details have changed
    const descriptionChanged = existingEvent.getDescription() !== `Type: ${event.type}\nEvent URL: ${event.url}\nHost Club: ${event.hostClub}\nRegion: ${event.region}`;
    const titleChanged = existingEvent.getTitle() !== eventTitle;
    const locationChanged = existingEvent.getLocation() !== event.hostClub;
    const startTimeChanged = existingEvent.getStartTime().getTime() !== startDate.getTime();
    const endTimeChanged = existingEvent.getEndTime().getTime() !== endDate.getTime();

    if (descriptionChanged || titleChanged || locationChanged || startTimeChanged || endTimeChanged) {
      // Update the event with new details
      Logger.log(`Updating event: ${eventTitle}`);
      
      existingEvent.setTitle(eventTitle);
      existingEvent.setLocation(event.hostClub);
      existingEvent.setDescription(`Type: ${event.type}\nEvent URL: ${event.url}\nHost Club: ${event.hostClub}\nRegion: ${event.region}`);
      existingEvent.setTime(startDate, endDate); // Update time (still all-day)
    } else {
      Logger.log(`Event with URL "${event.url}" already exists and is up-to-date.`);
    }
  } else {
    // No existing event, create a new one
    Logger.log(`Creating new event: ${eventTitle}`);
    
    calendar.createAllDayEvent(eventTitle, startDate, endDate, {
      location: event.hostClub,
      description: `Type: ${event.type}\nEvent URL: ${event.url}\nHost Club: ${event.hostClub}\nRegion: ${event.region}`
    });
  }
  
  Logger.log(`Processed event: ${eventTitle} from ${startDate} to ${endDate}`);
}
