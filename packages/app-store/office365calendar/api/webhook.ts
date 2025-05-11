import type { NextApiRequest, NextApiResponse } from "next";

import { HttpError } from "@calcom/lib/http-error";
import { defaultHandler, defaultResponder } from "@calcom/lib/server";
import prisma from "@calcom/prisma";
import { SelectedCalendarRepository } from "@calcom/features/calendar-cache/selectedCalendar.repository";
import { CalendarCache } from "@calcom/features/calendar-cache/calendar-cache";
import { getCalendar } from "@calcom/app-store/_utils/getCalendar";
import { getCalendarCredentials } from "@calcom/app-store/office365calendar/lib/getCalendarCredentials";
import logger from "@calcom/lib/logger";

const log = logger.getSubLogger({ prefix: ["office365calendar_webhook"] });

/**
 * This is the webhook endpoint that Microsoft Graph API will call when there are
 * changes to the calendars that we have subscribed to.
 *
 * @see https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks
 */
async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") {
    throw new HttpError({ statusCode: 405, message: "Method not allowed" });
  }

  // Validate the webhook token
  const validationToken = req.query.validationToken as string;
  if (validationToken) {
    // This is a subscription validation request
    log.debug("Received validation request", { validationToken });
    return res.status(200).send(validationToken);
  }

  // Verify the webhook request
  const clientState = req.headers["clientstate"] as string;
  if (clientState !== process.env.MICROSOFT_WEBHOOK_TOKEN) {
    log.error("Invalid client state token", { clientState });
    return res.status(401).json({ message: "Invalid client state token" });
  }

  // Parse the notification payload
  const body = req.body;
  if (!body.value || !Array.isArray(body.value)) {
    log.error("Invalid webhook payload", { body });
    return res.status(400).json({ message: "Invalid webhook payload" });
  }

  // Process each notification
  for (const notification of body.value) {
    try {
      await processNotification(notification);
    } catch (error) {
      log.error("Error processing notification", { error, notification });
    }
  }

  // Return a 202 Accepted response to acknowledge receipt of the notification
  return res.status(202).json({ message: "Notification received" });
}

/**
 * Process a single notification from Microsoft Graph API
 */
async function processNotification(notification: any) {
  const { subscriptionId, resourceData, changeType } = notification;
  
  if (!subscriptionId) {
    log.error("Missing subscriptionId in notification", { notification });
    return;
  }

  // Find the selected calendar associated with this subscription
  const selectedCalendars = await SelectedCalendarRepository.findMany({
    where: {
      microsoftSubscriptionId: subscriptionId,
    },
  });

  if (!selectedCalendars.length) {
    log.warn("No selected calendar found for subscription", { subscriptionId });
    return;
  }

  // Get the credential for the first selected calendar
  // All selected calendars with the same subscription should have the same credential
  const credentialId = selectedCalendars[0].credentialId;
  
  // Get the calendar service
  const calendarCredentials = await getCalendarCredentials(credentialId);
  if (!calendarCredentials) {
    log.error("No calendar credentials found", { credentialId });
    return;
  }

  const calendar = await getCalendar(calendarCredentials);
  if (!calendar) {
    log.error("Failed to get calendar", { credentialId });
    return;
  }

  // Get all integration calendars for this credential
  const integrationCalendars = await calendar.listCalendars();
  
  // Filter to only the calendars that match the selected calendars
  const selectedIntegrationCalendars = integrationCalendars.filter((cal) => 
    selectedCalendars.some((sc) => sc.externalId === cal.externalId)
  );

  // Delete the cache to force a refresh
  await prisma.calendarCache.deleteMany({
    where: {
      credentialId,
    },
  });

  // Fetch fresh availability data and update the cache
  if (calendar.fetchAvailabilityAndSetCache && selectedIntegrationCalendars.length > 0) {
    await calendar.fetchAvailabilityAndSetCache(selectedIntegrationCalendars);
    log.info("Updated calendar cache after webhook notification", { 
      subscriptionId, 
      changeType,
      calendarCount: selectedIntegrationCalendars.length 
    });
  }
}

export default defaultHandler({
  POST: Promise.resolve({ default: defaultResponder(handler) }),
});