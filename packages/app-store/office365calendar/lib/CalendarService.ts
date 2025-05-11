import type { Calendar as OfficeCalendar, User, Event, Subscription } from "@microsoft/microsoft-graph-types-beta";
import type { DefaultBodyType } from "msw";

import dayjs from "@calcom/dayjs";
import { CalendarCache } from "@calcom/features/calendar-cache/calendar-cache";
import { SelectedCalendarRepository } from "@calcom/features/calendar-cache/selectedCalendar.repository";
import { getLocation, getRichDescription } from "@calcom/lib/CalEventParser";
import {
  CalendarAppDelegationCredentialInvalidGrantError,
  CalendarAppDelegationCredentialConfigurationError,
} from "@calcom/lib/CalendarAppError";
import { uniqueBy } from "@calcom/lib/array";
import { handleErrorsJson, handleErrorsRaw } from "@calcom/lib/errors";
import { formatCalEvent } from "@calcom/lib/formatCalendarEvent";
import logger from "@calcom/lib/logger";
import { safeStringify } from "@calcom/lib/safeStringify";
import prisma from "@calcom/prisma";
import type {
  Calendar,
  CalendarEvent,
  EventBusyDate,
  IntegrationCalendar,
  NewCalendarEventType,
  SelectedCalendarEventTypeIds,
  BufferedBusyTime,
} from "@calcom/types/Calendar";
import type { CredentialForCalendarServiceWithTenantId } from "@calcom/types/Credential";
import { v4 as uuid } from "uuid";

// Import the functions from OutlookCalendarCache
import { 
  watchCalendar as watchCalendarFn, 
  unwatchCalendar as unwatchCalendarFn, 
  fetchAvailabilityAndSetCache as fetchAvailabilityAndSetCacheFn 
} from './OutlookCalendarCache';

import { OAuthManager } from "../../_utils/oauth/OAuthManager";
import { getTokenObjectFromCredential } from "../../_utils/oauth/getTokenObjectFromCredential";
import { oAuthManagerHelper } from "../../_utils/oauth/oAuthManagerHelper";
import metadata from "../_metadata";
import { getOfficeAppKeys } from "./getOfficeAppKeys";

interface IRequest {
  method: string;
  url: string;
  id: number;
}

interface ISettledResponse {
  id: string;
  status: number;
  headers: {
    "Retry-After": string;
    "Content-Type": string;
  };
  body: any;
}

interface IBatchResponse {
  responses: ISettledResponse[];
}

interface BodyValue {
  showAs: string;
  start: { dateTime: string };
  end: { dateTime: string };
}

export class Office365CalendarService implements Calendar {
  private url = "";
  private integrationName = "";
  private log: typeof logger;
  private auth: OAuthManager;
  private apiGraphUrl = "https://graph.microsoft.com/v1.0";
  private credential: CredentialForCalendarServiceWithTenantId;
  private azureUserId?: string;

  constructor(credential: CredentialForCalendarServiceWithTenantId) {
    this.integrationName = "office365_calendar";
    const tokenResponse = getTokenObjectFromCredential(credential);
    this.auth = new OAuthManager({
      credentialSyncVariables: oAuthManagerHelper.credentialSyncVariables,
      resourceOwner: {
        type: "user",
        id: credential.userId,
      },
      appSlug: metadata.slug,
      currentTokenObject: tokenResponse,
      fetchNewTokenObject: async ({ refreshToken }: { refreshToken: string | null }) => {
        const isDelegated = Boolean(credential?.delegatedTo);

        if (!isDelegated && !refreshToken) {
          return null;
        }

        const { client_id, client_secret } = await this.getAuthCredentials(isDelegated);

        const url = this.getAuthUrl(isDelegated, credential?.delegatedTo?.serviceAccountKey?.tenant_id);

        const bodyParams = {
          scope: isDelegated
            ? "https://graph.microsoft.com/.default"
            : "User.Read Calendars.Read Calendars.ReadWrite",
          client_id,
          client_secret,
          grant_type: isDelegated ? "client_credentials" : "refresh_token",
          ...(isDelegated ? {} : { refresh_token: refreshToken ?? "" }),
        };

        return fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams(bodyParams),
        });
      },
      isTokenObjectUnusable: async function () {
        // TODO: Implement this. As current implementation of CalendarService doesn't handle it. It hasn't been handled in the OAuthManager implementation as well.
        // This is a placeholder for future implementation.
        return null;
      },
      isAccessTokenUnusable: async function () {
        // TODO: Implement this
        return null;
      },

      invalidateTokenObject: () => oAuthManagerHelper.invalidateCredential(credential.id),
      expireAccessToken: () => oAuthManagerHelper.markTokenAsExpired(credential),
      updateTokenObject: (tokenObject) => {
        if (!Boolean(credential.delegatedTo)) {
          return oAuthManagerHelper.updateTokenObject({ tokenObject, credentialId: credential.id });
        }
        return Promise.resolve();
      },
    });
    this.credential = credential;
    this.log = logger.getSubLogger({ prefix: [`[[lib] ${this.integrationName}`] });
  }

  private getAuthUrl(delegatedTo: boolean, tenantId?: string): string {
    if (delegatedTo) {
      return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    }
    return "https://login.microsoftonline.com/common/oauth2/v2.0/token";
  }

  private async getAuthCredentials(isDelegated: boolean) {
    const appKeys = await getOfficeAppKeys();
    const { client_id, client_secret } = appKeys;

    if (isDelegated) {
      if (!client_id || !client_secret) {
        throw new CalendarAppDelegationCredentialConfigurationError();
      }
    }

    return {
      client_id,
      client_secret,
    };
  }

  private async getAzureUserId(credential: CredentialForCalendarServiceWithTenantId) {
    const isDelegated = Boolean(credential?.delegatedTo);
    if (isDelegated) {
      return credential.delegatedTo?.serviceAccountKey?.user_id;
    }
    return null;
  }

  // It would error if the delegation credential is not set up correctly
  async testDelegationCredentialSetup(): Promise<boolean> {
    const isDelegated = Boolean(this.credential?.delegatedTo);
    if (!isDelegated) {
      return true;
    }
    try {
      await this.getUserEndpoint();
      return true;
    } catch (error) {
      throw new CalendarAppDelegationCredentialInvalidGrantError();
    }
  }

  async getUserEndpoint(): Promise<string> {
    if (!this.azureUserId) {
      this.azureUserId = await this.getAzureUserId(this.credential);
    }
    if (this.azureUserId) {
      return `/users/${this.azureUserId}`;
    }
    return "/me";
  }

  /**
   * Subscribes to changes in an Outlook calendar using Microsoft Graph API
   * @param calendarId - The ID of the calendar to watch
   * @param eventTypeIds - The event type IDs to associate with this calendar
   */
  async watchCalendar({
    calendarId,
    eventTypeIds,
  }: {
    calendarId: string;
    eventTypeIds: SelectedCalendarEventTypeIds;
  }): Promise<string | undefined> {
    return watchCalendarFn({
      calendarId,
      eventTypeIds,
      auth: this.auth,
      apiGraphUrl: this.apiGraphUrl,
      credential: this.credential,
      log: this.log,
    });
  }

  /**
   * Unsubscribes from changes in an Outlook calendar
   * @param calendarId - The ID of the calendar to unwatch
   * @param eventTypeIds - The event type IDs to disassociate from this calendar
   */
  async unwatchCalendar({
    calendarId,
    eventTypeIds,
  }: {
    calendarId: string;
    eventTypeIds: SelectedCalendarEventTypeIds;
  }): Promise<void> {
    return unwatchCalendarFn({
      calendarId,
      eventTypeIds,
      auth: this.auth,
      apiGraphUrl: this.apiGraphUrl,
      credential: this.credential,
      log: this.log,
    });
  }
  
  /**
   * Fetches availability data and updates the cache
   * @param calendars - The calendars to fetch availability for
   */
  async fetchAvailabilityAndSetCache(calendars: IntegrationCalendar[]): Promise<any> {
    return fetchAvailabilityAndSetCacheFn({
      credential: this.credential,
      calendars,
      getAvailability: (options) => this.getAvailability(
        options.dateFrom.toISOString(), 
        options.dateTo.toISOString(), 
        options.selectedCalendars
      ),
      log: this.log,
    });
  }
  
  private translateEvent = (event: CalendarEvent) => {
    const office365Event = {
      subject: event.title,
      body: {
        contentType: "text",
        content: getRichDescription(event),
      },
      start: {
        dateTime: dayjs(event.startTime).tz(event.organizer.timeZone).format("YYYY-MM-DDTHH:mm:ss"),
        timeZone: event.organizer.timeZone,
      },
      end: {
        dateTime: dayjs(event.endTime).tz(event.organizer.timeZone).format("YYYY-MM-DDTHH:mm:ss"),
        timeZone: event.organizer.timeZone,
      },
      hideAttendees: !event.seatsPerTimeSlot ? false : !event.seatsShowAttendees,
      organizer: {
        emailAddress: {
          address: event.destinationCalendar
            ? event.destinationCalendar.find((cal) => cal.userId === event.organizer.id)?.externalId ??
              event.organizer.email
            : event.organizer.email,
          name: event.organizer.name,
        },
      },
      attendees: [
        ...event.attendees.map((attendee) => ({
          emailAddress: {
            address: attendee.email,
            name: attendee.name,
          },
          type: "required" as const,
        })),
        ...(event.team?.members
          ? event.team?.members
              .filter((member) => member.email !== this.credential.user?.email)
              .map((member) => {
                const destinationCalendar =
                  event.destinationCalendar &&
                  event.destinationCalendar.find((cal) => cal.userId === member.id);
                return {
                  emailAddress: {
                    address: destinationCalendar?.externalId ?? member.email,
                    name: member.name,
                  },
                  type: "required" as const,
                };
              })
          : []),
      ],
      location: event.location ? { displayName: getLocation(event) } : undefined,
    };
    if (event.hideCalendarEventDetails) {
      office365Event["sensitivity"] = "private";
    }
    return office365Event;
  };

  private fetcher = async (endpoint: string, init?: RequestInit | undefined) => {
    return this.auth.requestRaw({
      url: `${this.apiGraphUrl}${endpoint}`,
      options: {
        method: "get",
        ...init,
      },
    });
  };

  private fetchResponsesWithNextLink = async (
    settledResponses: ISettledResponse[]
  ): Promise<ISettledResponse[]> => {
    const alreadySuccess = [] as ISettledResponse[];
    const newLinkRequest = [] as IRequest[];
    settledResponses?.forEach((response) => {
      if (response.status === 200 && response.body["@odata.nextLink"] === undefined) {
        alreadySuccess.push(response);
      } else {
        const nextLinkUrl = response.body["@odata.nextLink"]
          ? String(response.body["@odata.nextLink"]).replace(this.apiGraphUrl, "")
          : "";
        if (nextLinkUrl) {
          // Saving link for later use
          newLinkRequest.push({
            id: Number(response.id),
            method: "GET",
            url: nextLinkUrl,
          });
        }
        delete response.body["@odata.nextLink"];
        // Pushing success body content
        alreadySuccess.push(response);
      }
    });

    if (newLinkRequest.length === 0) {
      return alreadySuccess;
    }

    const newResponse = await this.apiGraphBatchCall(newLinkRequest);
    let newResponseBody = await handleErrorsJson<IBatchResponse | string>(newResponse);

    if (typeof newResponseBody === "string") {
      newResponseBody = this.handleTextJsonResponseWithHtmlInBody(newResponseBody);
    }

    // Going recursive to fetch next link
    const newSettledResponses = await this.fetchResponsesWithNextLink(newResponseBody.responses);
    return [...alreadySuccess, ...newSettledResponses];
  };

  private fetchRequestWithRetryAfter = async (
    originalRequests: IRequest[],
    settledPromises: ISettledResponse[],
    maxRetries: number,
    retryCount = 0
  ): Promise<IBatchResponse> => {
    const getRandomness = () => Number(Math.random().toFixed(3));
    let retryAfterTimeout = 0;
    if (retryCount >= maxRetries) {
      return { responses: settledPromises };
    }
    const alreadySuccessRequest = [] as ISettledResponse[];
    const failedRequest = [] as IRequest[];
    settledPromises.forEach((item) => {
      if (item.status === 200) {
        alreadySuccessRequest.push(item);
      } else if (item.status === 429) {
        const newTimeout = Number(item.headers["Retry-After"]) * 1000 || 0;
        retryAfterTimeout = newTimeout > retryAfterTimeout ? newTimeout : retryAfterTimeout;
        failedRequest.push(originalRequests[Number(item.id)]);
      }
    });

    if (failedRequest.length === 0) {
      return { responses: alreadySuccessRequest };
    }

    // Await certain time from retry-after header
    await new Promise((r) => setTimeout(r, retryAfterTimeout + getRandomness()));

    const newResponses = await this.apiGraphBatchCall(failedRequest);
    let newResponseBody = await handleErrorsJson<IBatchResponse | string>(newResponses);
    if (typeof newResponseBody === "string") {
      newResponseBody = this.handleTextJsonResponseWithHtmlInBody(newResponseBody);
    }
    const retryAfter = !!newResponseBody?.responses && this.findRetryAfterResponse(newResponseBody.responses);

    if (retryAfter && newResponseBody.responses) {
      newResponseBody = await this.fetchRequestWithRetryAfter(
        failedRequest,
        newResponseBody.responses,
        maxRetries,
        retryCount + 1
      );
    }
    return { responses: [...alreadySuccessRequest, ...(newResponseBody?.responses || [])] };
  };

  private apiGraphBatchCall = async (requests: IRequest[]): Promise<Response> => {
    const response = await this.fetcher(`/$batch`, {
      method: "POST",
      body: JSON.stringify({ requests }),
    });

    return response;
  };

  private handleTextJsonResponseWithHtmlInBody = (response: string): IBatchResponse => {
    try {
      const parsedJson = JSON.parse(response);
      return parsedJson;
    } catch (error) {
      // Looking for html in body
      const openTag = '"body":<';
      const closeTag = "</html>";
      const htmlBeginning = response.indexOf(openTag) + openTag.length - 1;
      const htmlEnding = response.indexOf(closeTag) + closeTag.length + 2;
      const resultString = `${response.substring(0, htmlBeginning)} ""${response
        .substring(htmlEnding, response.length)}`;

      return JSON.parse(resultString);
    }
  };

  private findRetryAfterResponse = (response: ISettledResponse[]) => {
    const foundRetry = response.find((request: ISettledResponse) => request.status === 429);
    return !!foundRetry;
  };

  private processBusyTimes = (responses: ISettledResponse[]) => {
    return responses.reduce(
      (acc: BufferedBusyTime[], subResponse: { body: { value?: BodyValue[]; error?: Error[] } }) => {
        if (!subResponse.body?.value) return acc;
        return acc.concat(
          subResponse.body.value.reduce((acc: BufferedBusyTime[], evt: BodyValue) => {
            if (evt.showAs === "free" || evt.showAs === "workingElsewhere") return acc;
            return acc.concat({
              start: `${evt.start.dateTime}Z`,
              end: `${evt.end.dateTime}Z`,
            });
          }, [])
        );
      },
      []
    );
  };

  private handleErrorJsonOffice365Calendar = <Type>(response: Response): Promise<Type | string> => {
    if (response.headers.get("content-encoding") === "gzip") {
      return response.text();
    }

    if (response.status === 204) {
      return new Promise((resolve) => resolve({} as Type));
    }

    if (!response.ok && response.status < 200 && response.status >= 300) {
      response.json().then(console.log);
      throw Error(response.statusText);
    }

    return response.json();
  };

  async createEvent(event: CalendarEvent, credentialId: number): Promise<NewCalendarEventType> {
    const userEndpoint = await this.getUserEndpoint();
    const calendarId = event.destinationCalendar?.externalId;
    this.log.debug("Creating Office 365 Event", safeStringify({ userEndpoint, calendarId }));

    const response = await this.auth.request({
      method: "POST",
      url: `${this.apiGraphUrl}${userEndpoint}/calendars/${calendarId}/events`,
      body: this.translateEvent(event),
    });

    return handleErrorsJson<NewCalendarEventType>(response);
  }

  async updateEvent(
    uid: string,
    event: CalendarEvent
  ): Promise<NewCalendarEventType | NewCalendarEventType[]> {
    const userEndpoint = await this.getUserEndpoint();
    const calendarId = event.destinationCalendar?.externalId;

    const eventsToUpdate = await this.getEventsByICalUID(uid);

    const eventsToUpdatePromises = eventsToUpdate.map((eventToUpdate) => {
      return this.auth.request({
        method: "PATCH",
        url: `${this.apiGraphUrl}${userEndpoint}/calendars/${calendarId}/events/${eventToUpdate.id}`,
        body: this.translateEvent(event),
      });
    });

    const responses = await Promise.all(eventsToUpdatePromises);
    const responsePromises = responses.map((response) => handleErrorsJson<NewCalendarEventType>(response));
    return Promise.all(responsePromises);
  }

  async deleteEvent(uid: string): Promise<void> {
    const userEndpoint = await this.getUserEndpoint();
    const events = await this.getEventsByICalUID(uid);

    const eventsToDeletePromises = events.map((event) => {
      return this.auth.request({
        method: "DELETE",
        url: `${this.apiGraphUrl}${userEndpoint}/events/${event.id}`,
      });
    });

    await Promise.all(eventsToDeletePromises);
  }

  async getAvailability(
    dateFrom: string,
    dateTo: string,
    selectedCalendars: IntegrationCalendar[]
  ): Promise<EventBusyDate[]> {
    const userEndpoint = await this.getUserEndpoint();
    const dateFromParsed = new Date(dateFrom);
    const dateToParsed = new Date(dateTo);

    const filter = `?startDateTime=${encodeURIComponent(
      dateFromParsed.toISOString()
    )}&endDateTime=${encodeURIComponent(dateToParsed.toISOString())}`;

    try {
      // Check if there's a cached version of the availability data
      const cachedAvailability = await prisma.calendarCache.findFirst({
        where: {
          credentialId: this.credential.id,
          key: `availability:${this.credential.id}`,
          expiresAt: {
            gt: new Date(),
          },
        },
      });

      if (cachedAvailability) {
        this.log.debug("Using cached availability data", { credentialId: this.credential.id });
        return JSON.parse(cachedAvailability.value);
      }

      const selectedCalendarIds = selectedCalendars
        .filter((sc) => sc.integration === this.integrationName)
        .map((sc) => sc.externalId);
      if (selectedCalendarIds.length === 0 && selectedCalendars.length > 0) {
        // Only calendars of other integrations selected
        return Promise.resolve([]);
      }

      const calendarSelectionsForCalendarIds = selectedCalendarIds.map((calendarId, id) => {
        return {
          id,
          method: "GET",
          url: `${userEndpoint}/calendars/${calendarId}/calendarView${filter}`,
        };
      });

      const response = await this.apiGraphBatchCall(calendarSelectionsForCalendarIds);
      let responseBody = await handleErrorsJson<IBatchResponse | string>(response);
      if (typeof responseBody === "string") {
        responseBody = this.handleTextJsonResponseWithHtmlInBody(responseBody);
      }

      let responsesWithNextLink = await this.fetchResponsesWithNextLink(responseBody.responses);
      responsesWithNextLink = await this.fetchRequestWithRetryAfter(
        calendarSelectionsForCalendarIds,
        responsesWithNextLink,
        3
      ).then((val) => val.responses);
      return this.processBusyTimes(responsesWithNextLink);
    } catch (err) {
      this.log.error("Failed to get availability from Office 365", err);
      return Promise.reject([]);
    }
  }

  async listCalendars(event?: CalendarEvent): Promise<IntegrationCalendar[]> {
    try {
      const userEndpoint = await this.getUserEndpoint();
      const response = await this.fetcher(`${userEndpoint}/calendars`);
      let responseBody = await handleErrorsJson<{ value: OfficeCalendar[] } | string>(response);
      if (typeof responseBody === "string") {
        responseBody = JSON.parse(responseBody) as { value: OfficeCalendar[] };
      }
      return responseBody.value.map((cal: OfficeCalendar) => {
        const calendar: IntegrationCalendar = {
          externalId: cal.id ?? "No Id",
          integration: this.integrationName,
          name: cal.name ?? "No calendar name",
          primary: cal.isDefaultCalendar ?? false,
          readOnly: !cal.canEdit && true,
          email: cal.owner?.address ?? this.credential.user?.email ?? "",
        };
        return calendar;
      });
    } catch (err) {
      this.log.error("Failed to get calendars from Office 365", err);
      return Promise.reject([]);
    }
  }

  private async getEventsByICalUID(uid: string): Promise<{ id: string }[]> {
    const userEndpoint = await this.getUserEndpoint();
    const response = await this.fetcher(`${userEndpoint}/events?$filter=iCalUId eq '${uid}'`);
    const responseBody = await handleErrorsJson<{ value: { id: string }[] }>(response);
    return responseBody.value;
  }
}