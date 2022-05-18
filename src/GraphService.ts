import {
  Client,
  GraphRequestOptions,
  PageCollection,
  PageIterator,
} from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { endOfWeek, startOfWeek } from "date-fns";
import { zonedTimeToUtc } from "date-fns-tz";
import { User, Event } from "microsoft-graph";

let graphClient: Client | undefined = undefined;

function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  if (!graphClient) {
    graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });
  }

  return graphClient;
}

export async function getUserWeekCalendar(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  timeZone: string
): Promise<Event[]> {
  ensureClient(authProvider);

  // Generate startDateTime and endDateTime query params
  // to display a 7-day window
  const now = new Date();
  const startDateTime = zonedTimeToUtc(
    startOfWeek(now),
    timeZone
  ).toISOString();
  const endDateTime = zonedTimeToUtc(endOfWeek(now), timeZone).toISOString();

  // GET /me/calendarview?startDateTime=''&endDateTime=''
  // &$select=subject,organizer,start,end
  // &$orderby=start/dateTime
  // &$top=50
  var response: PageCollection = await graphClient!
    .api("/me/calendarview")
    .header("Prefer", `outlook.timezone="${timeZone}"`)
    .query({ startDateTime: startDateTime, endDateTime: endDateTime })
    .select("subject,organizer,start,end")
    .orderby("start/dateTime")
    .top(25)
    .get();

  if (response["@odata.nextLink"]) {
    // Presence of the nextLink property indicates more results are available
    // Use a page iterator to get all results
    var events: Event[] = [];

    // Must include the time zone header in page
    // requests too
    var options: GraphRequestOptions = {
      headers: { Prefer: `outlook.timezone="${timeZone}"` },
    };

    var pageIterator = new PageIterator(
      graphClient!,
      response,
      (event) => {
        events.push(event);
        return true;
      },
      options
    );

    await pageIterator.iterate();

    return events;
  } else {
    return response.value;
  }
}

export async function getUser(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider
): Promise<User> {
  ensureClient(authProvider);

  // Return the /me API endpoint result as a User object
  const user: User = await graphClient!
    .api("/me")
    // Only retrieve the specific fields needed
    .select("displayName,mail,mailboxSettings,userPrincipalName")
    .get();

  return user;
}

export async function createEvent(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  newEvent: Event
): Promise<Event> {
  ensureClient(authProvider);

  // POST /me/events
  // JSON representation of the new event is sent in the
  // request body
  return await graphClient!.api("/me/events").post(newEvent);
}

export async function searchDocs(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  siteId: string,
  query: string
): Promise<any> {
  ensureClient(authProvider);

  const matchedDocs = await graphClient!
    .api(`/sites/${siteId}/drive/root/search(q='${query}')`)
    .get();

  return matchedDocs.value;
}

export async function getDoc(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  siteId: string,
  id: string
): Promise<any> {
  ensureClient(authProvider);

  const matchedDoc = await graphClient!
    .api(`/sites/${siteId}/drive/items/${id}`)
    .get();

  return matchedDoc;
}

export async function getDocVersions(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  siteId: string,
  id: string
): Promise<any> {
  ensureClient(authProvider);

  const docVersions = await graphClient!
    .api(`/sites/${siteId}/drive/items/${id}/versions`)
    .get();

  return docVersions.value;
}

export async function uploadDocument(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  siteId: string,
  filename: string,
  fileToUpload: any
) {
  ensureClient(authProvider);
  const uploadURL = `/sites/${siteId}/drive/root:/${filename}:/content`;
  const response = await graphClient!.api(uploadURL).put(fileToUpload);
  return response;
}

export async function updateDoc(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  id: string,
  siteId: string,
  body: any
): Promise<any> {
  ensureClient(authProvider);
  const updatedDoc = await graphClient!
    .api(`/sites/${siteId}/drive/items/${id}`)
    .update(body);

  return updatedDoc;
}

export async function deleteDoc(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  siteId: string,
  id: string
): Promise<any> {
  ensureClient(authProvider);
  const deletedDoc = await graphClient!
    .api(`/sites/${siteId}/drive/items/${id}`)
    .delete();

  return deletedDoc;
}

export async function searchSites(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  query: string
): Promise<any> {
  ensureClient(authProvider);
  const matchedSites = await graphClient!.api(`/sites?search=${query}`).get();
  return matchedSites.value;
}

export async function createGroup(
  authProvider: AuthCodeMSALBrowserAuthenticationProvider,
  displayName: string,
  description: string,
  mailNickName: string
): Promise<any> {
  ensureClient(authProvider);
  const newGroup = await graphClient!.api("/groups").post({
    displayName: displayName,
    description: description,
    groupTypes: ["Unified"],
    mailEnabled: true,
    mailNickname: mailNickName,
    securityEnabled: false,
  });
  return newGroup;
}
