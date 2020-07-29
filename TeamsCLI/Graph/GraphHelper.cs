using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace TeamsCLI
{
    class GraphHelper
    {
        private static GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

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
                        e.Organizer,
                        e.Start,
                        e.End
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

        public static async Task<IEnumerable<Chat>> GetChatsAsync()
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
                var resultPage = await graphClient.Chats[chatId].Messages.Request()
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

                return resultPage.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting chat {chatId} messages: {ex.Message}");
                return null;
            }
        }
    }
}
