import type { SelectedCalendarEventTypeIds } from "@calcom/types/Calendar";
import type { CredentialForCalendarServiceWithTenantId } from "@calcom/types/Credential";
import type { IntegrationCalendar } from "@calcom/types/Calendar";
import type { OAuthManager } from "../../_utils/oauth/OAuthManager";
import { safeStringify } from "@calcom/lib/safeStringify";
import prisma from "@calcom/prisma";
import logger from "@calcom/lib/logger";

// Use the WEBAPP_URL from the environment variables
const WEBAPP_URL = process.env.WEBAPP_URL || process.env.NEXT_PUBLIC_WEBAPP_URL || "http://localhost:3000";

/**
 * Subscribes to changes in an Outlook calendar using Microsoft Graph API
 */
export async function watchCalendar({
  calendarId,
  eventTypeIds,
  auth,
  apiGraphUrl,
  credential,
  log,
}: {
  calendarId: string;
  eventTypeIds: SelectedCalendarEventTypeIds;
  auth: OAuthManager;
  apiGraphUrl: string;
  credential: CredentialForCalendarServiceWithTenantId;
  log: typeof logger;
}): Promise<string | undefined> {
  log.debug("watchCalendar", safeStringify({ calendarId, eventTypeIds }));
  
  if (!process.env.MICROSOFT_WEBHOOK_TOKEN) {
    log.warn("MICROSOFT_WEBHOOK_TOKEN is not set, skipping watching calendar");
    return;
  }

  // Check if this calendar is already being watched
  const existingSubscription = await prisma.selectedCalendar.findFirst({
    where: {
      externalId: calendarId,
      microsoftSubscriptionId: { not: null },
    },
  });

  if (existingSubscription?.microsoftSubscriptionId) {
    log.debug(
      "Calendar already being watched",
      safeStringify({ calendarId, subscriptionId: existingSubscription.microsoftSubscriptionId })
    );

    // Update the event type IDs for this calendar
    await prisma.selectedCalendar.updateMany({
      where: {
        externalId: calendarId,
        userId: credential.userId,
      },
      data: {
        eventTypeId: eventTypeIds.eventTypeId,
      },
    });

    return existingSubscription.microsoftSubscriptionId;
  }

  try {
    // Create a new subscription
    const webhookUrl = `${WEBAPP_URL}/api/integrations/office365calendar/webhook`;
    
    // Subscription expires after 3 days (maximum allowed by Microsoft Graph API)
    const expirationDateTime = new Date();
    expirationDateTime.setDate(expirationDateTime.getDate() + 3);

    const response = await auth.request({
      method: "POST",
      url: `${apiGraphUrl}/subscriptions`,
      body: {
        changeType: "created,updated,deleted",
        notificationUrl: webhookUrl,
        resource: `/me/calendars/${calendarId}/events`,
        expirationDateTime: expirationDateTime.toISOString(),
        clientState: process.env.MICROSOFT_WEBHOOK_TOKEN,
      },
    });

    if (!response.id) {
      throw new Error("Failed to create subscription");
    }

    // Update the selected calendar with the subscription ID
    await prisma.selectedCalendar.updateMany({
      where: {
        externalId: calendarId,
        userId: credential.userId,
      },
      data: {
        microsoftSubscriptionId: response.id,
        eventTypeId: eventTypeIds.eventTypeId,
      },
    });

    log.debug(
      "Created subscription",
      safeStringify({ calendarId, subscriptionId: response.id })
    );

    return response.id;
  } catch (error) {
    log.error("Error creating subscription", safeStringify({ calendarId, error }));
    throw error;
  }
}

/**
 * Unsubscribes from changes in an Outlook calendar
 */
export async function unwatchCalendar({
  calendarId,
  eventTypeIds,
  auth,
  apiGraphUrl,
  credential,
  log,
}: {
  calendarId: string;
  eventTypeIds: SelectedCalendarEventTypeIds;
  auth: OAuthManager;
  apiGraphUrl: string;
  credential: CredentialForCalendarServiceWithTenantId;
  log: typeof logger;
}): Promise<void> {
  log.debug("unwatchCalendar", safeStringify({ calendarId, eventTypeIds }));

  // Find all selected calendars with this external ID
  const allSelectedCalendars = await prisma.selectedCalendar.findMany({
    where: {
      externalId: calendarId,
      userId: credential.userId,
    },
  });

  // Find the subscription ID for this calendar
  const subscriptionId = allSelectedCalendars[0]?.microsoftSubscriptionId;
  if (!subscriptionId) {
    log.debug("No subscription found for calendar", safeStringify({ calendarId }));
    return;
  }

  // Check if there are other event types still using this calendar
  const eventTypeIdsToBeUnwatched = Array.isArray(eventTypeIds.eventTypeId)
    ? eventTypeIds.eventTypeId
    : [eventTypeIds.eventTypeId];

  const calendarsWithSameExternalIdThatAreBeingWatched = allSelectedCalendars.filter(
    (sc) => sc.microsoftSubscriptionId === subscriptionId
  );

  const calendarsWithSameExternalIdToBeStillWatched = calendarsWithSameExternalIdThatAreBeingWatched.filter(
    (sc) => !eventTypeIdsToBeUnwatched.includes(sc.eventTypeId)
  );

  // If there are still other event types using this calendar, just update the event type ID
  if (calendarsWithSameExternalIdToBeStillWatched.length > 0) {
    log.debug(
      "Calendar still being watched by other event types",
      safeStringify({ calendarId, remainingEventTypes: calendarsWithSameExternalIdToBeStillWatched.length })
    );

    // Update the selected calendar to remove the event type ID
    await prisma.selectedCalendar.deleteMany({
      where: {
        externalId: calendarId,
        userId: credential.userId,
        eventTypeId: { in: eventTypeIdsToBeUnwatched },
      },
    });

    return;
  }

  try {
    // Delete the subscription
    await auth.request({
      method: "DELETE",
      url: `${apiGraphUrl}/subscriptions/${subscriptionId}`,
    });

    // Update all selected calendars to remove the subscription ID
    await prisma.selectedCalendar.updateMany({
      where: {
        externalId: calendarId,
        userId: credential.userId,
      },
      data: {
        microsoftSubscriptionId: null,
      },
    });

    log.debug("Deleted subscription", safeStringify({ calendarId, subscriptionId }));
  } catch (error) {
    log.error("Error deleting subscription", safeStringify({ calendarId, subscriptionId, error }));
    
    // Even if the API call fails, remove the subscription ID from the database
    // This could happen if the subscription was already deleted on Microsoft's side
    await prisma.selectedCalendar.updateMany({
      where: {
        externalId: calendarId,
        userId: credential.userId,
      },
      data: {
        microsoftSubscriptionId: null,
      },
    });
  }
}

/**
 * Fetches availability data and updates the cache
 */
export async function fetchAvailabilityAndSetCache({
  credential,
  calendars,
  getAvailability,
  log,
}: {
  credential: CredentialForCalendarServiceWithTenantId;
  calendars: IntegrationCalendar[];
  getAvailability: (options: { dateFrom: Date; dateTo: Date; selectedCalendars: IntegrationCalendar[] }) => Promise<any>;
  log: typeof logger;
}): Promise<any> {
  const startTime = new Date();
  startTime.setDate(startTime.getDate() - 7); // Get events from the last week
  
  const endTime = new Date();
  endTime.setDate(endTime.getDate() + 30); // Get events for the next 30 days
  
  try {
    const availability = await getAvailability({
      dateFrom: startTime,
      dateTo: endTime,
      selectedCalendars: calendars,
    });
    
    // Store the availability data in the cache
    await prisma.calendarCache.create({
      data: {
        credentialId: credential.id,
        key: `availability:${credential.id}`,
        value: JSON.stringify(availability),
        expiresAt: new Date(Date.now() + 15 * 60 * 1000), // Cache for 15 minutes
      },
    });
    
    log.debug("Updated calendar cache", { credentialId: credential.id });
    return availability;
  } catch (error) {
    log.error("Error fetching availability for cache", { error });
    throw error;
  }
}