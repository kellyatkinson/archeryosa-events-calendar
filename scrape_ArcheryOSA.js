// This script scrapes ArcheryOSA upcoming event details and inserts them into a Google Calendar.
// It is suitable for daily use.
// Event updates are reflected when the script runs.
// Kelly Atkinson, 2025

const CONFIG = {
  CAL_ID: PropertiesService.getScriptProperties().getProperty('CALENDAR_ID'),
  BASE_URLS: [
    'https://archeryosa.com',
    'https://www.archeryosa.com'
  ],
  EVENT_LISTING_PATHS: ['/events']
};

const DEFAULT_FETCH_OPTIONS = {
  muteHttpExceptions: true,
  followRedirects: true,
  headers: {
    'User-Agent': 'Mozilla/5.0 (compatible; ArcheryOSAEventsBot/1.0; +https://script.google.com)'
  }
};

const FETCH_RETRY_CONFIG = {
  maxAttempts: 3,
  initialDelayMs: 1000
};

function fetchWithRetry(url, options) {
  const fetchOptions = Object.assign({}, DEFAULT_FETCH_OPTIONS, options || {});
  let lastError = null;

  for (let attempt = 0; attempt < FETCH_RETRY_CONFIG.maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, fetchOptions);
      const statusCode = response.getResponseCode();

      if (statusCode >= 200 && statusCode < 300) {
        return response;
      }

      const snippet = response.getContentText().slice(0, 200);
      lastError = new Error(`HTTP ${statusCode}: ${snippet}`);
    } catch (error) {
      lastError = error;
    }

    if (attempt < FETCH_RETRY_CONFIG.maxAttempts - 1) {
      const backoff = Math.pow(2, attempt) * FETCH_RETRY_CONFIG.initialDelayMs;
      Utilities.sleep(backoff);
    }
  }

  throw new Error(`Failed to fetch ${url}. ${lastError ? lastError.message : 'Unknown error'}`);
}

function fetchEventListingPage() {
  const errors = [];

  for (let i = 0; i < CONFIG.BASE_URLS.length; i++) {
    const url = CONFIG.BASE_URLS[i];
    for (let j = 0; j < CONFIG.EVENT_LISTING_PATHS.length; j++) {
      const path = CONFIG.EVENT_LISTING_PATHS[j];
      const fullUrl = url + path;

      try {
        const response = fetchWithRetry(fullUrl);
        Logger.log(`Fetched event listing from ${fullUrl}`);
        return { baseUrl: url, html: response.getContentText() };
      } catch (error) {
        const message = `${fullUrl}: ${error.message}`;
        errors.push(message);
        Logger.log(`Failed to fetch ${message}`);
      }
    }
  }

  throw new Error(`Unable to fetch event listing page. Attempts: ${errors.join('; ')}`);
}

function automateEventCreation() {
  try {
    const { baseUrl, html } = fetchEventListingPage();
    const events = parseEventData(html, baseUrl); // Parse the events from the table

    events.forEach(event => {
      createOrUpdateCalendarEvent(event); // Create Google Calendar events for each
    });
  } catch (error) {
    Logger.log(`automateEventCreation failed: ${error.message}`);
    throw error;
  }
}

function fetchEventDetails(eventUrl) {
  const response = fetchWithRetry(eventUrl); // Fetch the event details page
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

function parseEventData(html, baseUrl) {
  const events = [];

  // Regex to capture the event URL, name, date, region, and type from the main table
  const tableRowRegex = /<tr>\s*<th scope="row">\s*<a href="([^"]+)">\s*([^<]+)<\/a>\s*<\/th>\s*<td>([^<]+)<\/td>\s*<td>([^<]+)<\/td>\s*<td>([^<]+)<\/td>\s*<\/tr>/g;
  
  let match;
  while ((match = tableRowRegex.exec(html)) !== null) {
    const eventHref = match[1];
    const sanitizedBase = baseUrl.replace(/\/$/, '');
    const normalizedPath = eventHref.startsWith('/') ? eventHref : `/${eventHref}`;
    const eventUrl = eventHref.match(/^https?:\/\//i)
      ? eventHref
      : `${sanitizedBase}${normalizedPath}`;  // Full event URL
    const eventName = match[2].trim();  // Event name
    const eventDate = match[3].trim();  // Event start date (e.g., "16 Nov")
    const eventRegion = match[4].trim();  // Event region
    const eventType = match[5].trim();  // Event type

    // Fetch additional details from the event page
    let details = {};
    try {
      details = fetchEventDetails(eventUrl);
    } catch (error) {
      Logger.log(`Failed to fetch event details for ${eventUrl}: ${error.message}`);
    }
    const startDate = details.startDate || eventDate;  // Use fetched start date if available
    const endDate = details.endDate || startDate;  // Use fetched end date if available
    let hostClub = details.hostClub || '';  // Use fetched host club if available
    if (hostClub === 'Unknown Club') {
      hostClub = '';
    }

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

function normalizeKeyPart(value) {
  return (value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function generateEventKey(event, startDate) {
  const timezone = Session.getScriptTimeZone() || 'Etc/GMT';
  const normalizedDate = Utilities.formatDate(startDate, timezone, 'yyyy-MM-dd');
  const normalizedHost = normalizeKeyPart(event.hostClub);
  const normalizedRegion = normalizeKeyPart(event.region);
  const normalizedType = normalizeKeyPart(event.type);

  return [normalizedDate, normalizedHost, normalizedRegion, normalizedType].join('|');
}

function buildEventDescription(event, eventKey) {
  return [
    `Type: ${event.type}`,
    `Event URL: ${event.url}`,
    `Host Club: ${event.hostClub}`,
    `Region: ${event.region}`,
    `Event Key: ${eventKey}`
  ].join('\n');
}

function extractDescriptionMetadata(description) {
  const metadata = {};

  if (!description) {
    return metadata;
  }

  const lines = description.split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const separatorIndex = line.indexOf(':');

    if (separatorIndex === -1) {
      continue;
    }

    const key = line.slice(0, separatorIndex).trim();
    const value = line.slice(separatorIndex + 1).trim();

    if (key) {
      metadata[key] = value;
    }
  }

  return metadata;
}

function findExistingCalendarEvent(calendar, startDate, event, eventKey) {
  const eventsOnDay = calendar.getEventsForDay(startDate);
  const normalizedHost = normalizeKeyPart(event.hostClub);
  const normalizedRegion = normalizeKeyPart(event.region);
  const normalizedType = normalizeKeyPart(event.type);

  for (let i = 0; i < eventsOnDay.length; i++) {
    const candidate = eventsOnDay[i];
    const description = candidate.getDescription() || '';
    const metadata = extractDescriptionMetadata(description);

    if (metadata['Event Key'] === eventKey) {
      return candidate;
    }

    if (metadata['Event URL'] === event.url) {
      return candidate;
    }

    const metadataHost = normalizeKeyPart(metadata['Host Club']);
    const metadataRegion = normalizeKeyPart(metadata['Region']);
    const metadataType = normalizeKeyPart(metadata['Type']);

    const hostMatches = normalizedHost && metadataHost === normalizedHost;
    const regionMatches = normalizedRegion && metadataRegion === normalizedRegion;
    const typeMatches = normalizedType && metadataType === normalizedType;

    // Require at least two matching attributes to avoid false positives.
    if (
      (hostMatches && regionMatches) ||
      (hostMatches && typeMatches) ||
      (regionMatches && typeMatches)
    ) {
      return candidate;
    }
  }

  return null;
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

  const eventKey = generateEventKey(event, startDate);
  const eventDescription = buildEventDescription(event, eventKey);
  const existingEvent = findExistingCalendarEvent(calendar, startDate, event, eventKey);

  let eventTitle;

  // Modify event title based on the hostClub and event type
  if (!event.hostClub || event.type === "League") {
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

  const eventLocation = event.hostClub || '';

  if (existingEvent) {
    // Update the existing event
    // Check if any details have changed
    const descriptionChanged = existingEvent.getDescription() !== eventDescription;
    const titleChanged = existingEvent.getTitle() !== eventTitle;
    const locationChanged = (existingEvent.getLocation() || '') !== eventLocation;
    const startTimeChanged = existingEvent.getStartTime().getTime() !== startDate.getTime();
    const endTimeChanged = existingEvent.getEndTime().getTime() !== endDate.getTime();

    if (descriptionChanged || titleChanged || locationChanged || startTimeChanged || endTimeChanged) {
      // Update the event with new details
      Logger.log(`Updating event: ${eventTitle}`);
      
      existingEvent.setTitle(eventTitle);
      existingEvent.setLocation(eventLocation);
      existingEvent.setDescription(eventDescription);
      existingEvent.setTime(startDate, endDate); // Update time (still all-day)
    } else {
      Logger.log(`Event with URL "${event.url}" already exists and is up-to-date.`);
    }
  } else {
    // No existing event, create a new one
    Logger.log(`Creating new event: ${eventTitle}`);
    
    calendar.createAllDayEvent(eventTitle, startDate, endDate, {
      location: eventLocation,
      description: eventDescription
    });
  }
  
  Logger.log(`Processed event: ${eventTitle} from ${startDate} to ${endDate}`);
}
