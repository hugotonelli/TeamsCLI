using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeamsCLI
{
    public class Meeting
    {
        public Event Event { get; set; }
        public Reminder Reminder { get; set; }
    }

    class GraphHelper
    {
        private static GraphServiceClient graphClient;
        private static IChatMessagesCollectionRequest chatMessagesNextPageRequest;
        private static readonly int chatMessagesPageSize = 50;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

        #region Me

        public static async Task<User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await graphClient.Me.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        #endregion


        #region Events

        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            try
            {
                // GET /me/events
                var resultPage = await graphClient.Me.Events.Request()
                    // Only return the fields used by the application
                    .Select(e => new
                    {
                        e.Subject,
                        e.Attendees,
                        e.Organizer,
                        e.Start,
                        e.End,
                    })
                    // Sort results by when they were created, newest first
                    .OrderBy("createdDateTime DESC")
                    .GetAsync();

                return resultPage.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<Event>> GetCalendarItemsAsync()
        {
            try
            {
                var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("startDateTime", DateTime.UtcNow.ToString("s")),
                    new QueryOption("endDateTime", "2021-01-07T19:00:00-08:00")
                };

                // GET /me/calendar
                var resultPage = await graphClient.Me.Calendar.CalendarView.Request(queryOptions)
                    .GetAsync();

                return resultPage.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting calendar: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<Calendar>> GetCalendarsAsync()
        {
            try
            {
                var calendars = await graphClient.Me.Calendars.Request()
                    .Select(c => new
                    {
                        c.Events,
                        c.Id,
                        c.Name,
                        c.Owner
                    })
                    .GetAsync();

                return calendars.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting calendars: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<Event>> GetEventsInAllCalendarsAsync()
        {
            try
            {
                var allCalendars = await GetCalendarsAsync();
                List<Event> allEvents = new List<Event>();

                foreach(var cal in allCalendars)
                {
                    if (cal.Events != null)
                    {
                        allEvents.AddRange(cal.Events);
                    }
                    else
                    {
                        var events = await graphClient.Me.Calendars[cal.Id].Events.Request().GetAsync();
                        allEvents.AddRange(events.CurrentPage);
                    }
                }

                return allEvents;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting all events for all calendars : {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<Reminder>> GetRemindersAsync()
        {
            try
            {
                var reminderView = await graphClient.Me
                    .ReminderView(
                        DateTime.UtcNow.ToString("s"),
                        DateTime.UtcNow.AddMonths(2).ToString("s"))
                    .Request()
                    .GetAsync();

                return reminderView.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting event reminders: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<Meeting>> GetEventsWithRemindersAsync()
        {
            try
            {
                var meetings = new List<Meeting>();
                var reminders = await GetRemindersAsync();

                foreach(var reminder in reminders)
                {
                    var meetingEvent = await graphClient.Me.Events[reminder.EventId].Request().GetAsync();
                    meetings.Add(new Meeting()
                    {
                        Reminder = reminder,
                        Event = meetingEvent,
                    });
                };

                return meetings;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events with reminders: {ex.Message}");
                throw;
            }
        }

        #endregion


        #region Chats

        public static async Task<IEnumerable<Chat>> GetChats()
        {
            try
            {
                // GET /me/chats
                var resultPage = await graphClient.Me.Chats.Request()
                    .Select(c => new
                    {
                        c.Id,
                        c.Topic,
                        //c.CreatedDateTime,
                        //c.LastUpdatedDateTime,
                        //c.Members
                    })
                    .GetAsync();

                /* // Too much overhead results in forbidden requests error.
                if (resultPage.CurrentPage != null)
                {
                    foreach (var chat in resultPage.CurrentPage)
                    {
                        if (chat.Members == null)
                        {
                            chat.Members = await graphClient.Chats[chat.Id].Members.Request().GetAsync();
                        }
                    }
                }
                */

                return resultPage.CurrentPage;

            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting chats: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<ConversationMember>> GetChatMembers(string chatId)
        {
            try
            {
                var resultPage = await graphClient.Chats[chatId].Members.Request()
                    .Select(m => new {
                        m.Id,
                        m.DisplayName,
                    })
                    .GetAsync();

                return resultPage.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting members for chat {chatId}: {ex.Message}");
                return null;
            }
        }

        public static async Task<Chat> GetSingleChatAsync(string chatId)
        {
            try
            {
                var result = await graphClient.Chats[chatId].Request()
                    .GetAsync();

                return result;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting chat {chatId}: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<ChatMessage>> GetChatMessages(string chatId)
        {
            try
            {
                chatMessagesNextPageRequest = null;

                var resultPage = await graphClient.Chats[chatId].Messages.Request()
                    .Top(chatMessagesPageSize)
                    //.Select(msg => new
                    //{
                    //    msg.Id,
                    //    msg.Attachments,
                    //    msg.Body,
                    //    msg.CreatedDateTime,
                    //    msg.From,
                    //    msg.Importance,
                    //    msg.Mentions,
                    //    msg.Summary,
                    //})
                    .GetAsync();

                chatMessagesNextPageRequest = resultPage.NextPageRequest;

                return resultPage.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting chat {chatId} messages: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<ChatMessage>> GetChatMessagesNextPage()
        {
            try
            {
                if (chatMessagesNextPageRequest != null)
                {
                    var resultPage = await chatMessagesNextPageRequest
                        .Top(chatMessagesPageSize)
                        .GetAsync();

                    chatMessagesNextPageRequest = resultPage.NextPageRequest;

                    return resultPage.CurrentPage;
                }

                return new List<ChatMessage>();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting next page of chat messages: {ex.Message}");
                return null;
            }
        }

        public static async Task<ChatMessage> PostChatMessage(string chatId, string message)
        {
            try
            {
                var chatMessage = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        Content = message
                    }
                };

                return await graphClient.Chats[chatId].Messages.Request()
                    .AddAsync(chatMessage);
                //return null;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error posting message to chat {chatId}: {ex.Message}");
                return null;
            }
        }

        #endregion
    }
}
