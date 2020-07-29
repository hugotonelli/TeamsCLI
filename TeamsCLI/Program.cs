using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace TeamsCLI
{
    class Program
    {
        const int maxChatsToDisplay = 25;
        static string testChatId = "";
        static string cancelString = "0";
        static string chatIdRegexPattern = @"^\d+:[\d\w-_]+@(([\d\w-_]*)\.)+[\d\w-_]*$";

        static void Main(string[] args)
        {
            Console.WriteLine("Teams CLI\n");

            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var tenantId = appConfig["tenantId"];
            var scopes = scopesString.Split(';');

            // Initialize the auth provider with values from appsettings.json
            var authProvider = new DeviceCodeAuthProvider(appId, scopes, tenantId);

            // Request a token to sign in the user
            var accessToken = authProvider.GetAccessToken().Result;

            // Initialize Graph client
            GraphHelper.Initialize(authProvider);

            // Get signed-in user
            var user = GraphHelper.GetMeAsync().Result;
            Console.WriteLine($"Welcome {user.DisplayName}!\n");

            int choice = -1;

            while (choice != 0)
            {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Display access token");
                Console.WriteLine("2. List calendar events");
                Console.WriteLine($"3. List top {maxChatsToDisplay} chats");
                Console.WriteLine("4. Chat info");
                Console.WriteLine("5. Chat messages");

                try
                {
                    choice = int.Parse(Console.ReadLine());
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }

                switch(choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye");
                        break;
                    case 1:
                        // Display access token
                        Console.WriteLine($"Access token: {accessToken}\n");
                        break;
                    case 2:
                        // List the calendar
                        ListCalendarEvents();
                        break;
                    case 3:
                        // List chats
                        ListChats();
                        break;
                    case 4:
                        GetChatId();
                        if (testChatId != cancelString)
                        {
                            // Show chat info for chatId
                            ChatInfo(testChatId);
                        }
                        break;
                    case 5:
                        GetChatId();
                        if (testChatId != cancelString)
                        {
                            // Show messages for chatId
                            ChatMessages(testChatId);
                        }
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }

        static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]) ||
                string.IsNullOrEmpty(appConfig["tenantId"]))
            {
                return null;
            }

            return appConfig;
        }

        static string FormatDateTimeTimeZone(Microsoft.Graph.DateTimeTimeZone value)
        {
            // Get the timezone specified in the Graph value
            var timeZone = TimeZoneInfo.FindSystemTimeZoneById(value.TimeZone);
            // Parse the date/time string from Graph into a DateTime
            var dateTime = DateTime.Parse(value.DateTime);

            // Create a DateTimeOffset in the specific timezone indicated by Graph
            var dateTimeWithTZ = new DateTimeOffset(dateTime, timeZone.BaseUtcOffset)
                .ToLocalTime();

            return dateTimeWithTZ.ToString("g");
        }

        static void ListCalendarEvents()
        {
            var events = GraphHelper.GetEventsAsync().Result;

            if (events == null)
            {
                return;
            }

            Console.WriteLine("Events: ");

            foreach (var calendarEvent in events)
            {
                Console.WriteLine($"Subject: {calendarEvent.Subject}");
                Console.WriteLine($"  Organizer: {calendarEvent.Organizer.EmailAddress.Name}");
                Console.WriteLine($"  Start: {FormatDateTimeTimeZone(calendarEvent.Start)}");
                Console.WriteLine($"  End: {FormatDateTimeTimeZone(calendarEvent.End)}");
            }
        }

        static void ListChats()
        {
            var chats = GraphHelper.GetChatsAsync().Result;

            if (chats == null)
            {
                Console.WriteLine("No chats found.");
                return;
            }

            chats = chats.Take(maxChatsToDisplay);

            Console.WriteLine("Chats: \n");

            foreach (var chat in chats)
            {
                Console.WriteLine($"Chat: {chat.Id}");
                if (String.IsNullOrEmpty(chat.Topic))
                {
                    ListChatMembers(chat.Id);
                }
                else
                {
                    Console.WriteLine($"  Topic: {chat.Topic}");
                }
                Console.WriteLine("");
            }
        }

        static void ListChatMembers(string chatId)
        {
            var chatMembers = GraphHelper.GetChatMembers(chatId).Result;

            if (chatMembers != null)
            {
                Console.WriteLine($"  Members: {String.Join(", ", chatMembers.Select(m => m.DisplayName))}");
            }
        }

        static void ChatInfo(string chatId)
        {
            var chat = GraphHelper.GetSingleChatAsync(chatId).Result;

            if (chat == null)
            {
                Console.WriteLine("Chat not found");
                return;
            }

            Console.WriteLine($"Chat: {chatId}");
            if (String.IsNullOrEmpty(chat.Topic))
            {
                ListChatMembers(chat.Id);
            }
            else
            {
                Console.WriteLine($"  Topic: {chat.Topic}");
            }
            Console.WriteLine("");
        }

        static void ChatMessages(string chatId)
        {
            var chatMessages = GraphHelper.GetChatMessages(chatId).Result.Take(maxChatsToDisplay).Reverse();

            if (chatMessages == null)
            {
                Console.WriteLine("Chat messages not found");
                return;
            }

            Console.WriteLine("Chat messages:");

            foreach (var msg in chatMessages)
            {
                // Console.WriteLine($"Id: {msg.Id}");
                Console.WriteLine($"From: {msg.From.User.DisplayName} At: {msg.CreatedDateTime.Value.ToLocalTime()}");
                //Console.WriteLine(msg.CreatedDateTime);
                if (msg.Attachments.Any()) Console.WriteLine("Has attachments");
                if (msg.Importance.HasValue && msg.Importance.Value != ChatMessageImportance.Normal) Console.WriteLine("[IMPORTANT!!!]");
                Console.WriteLine(msg.Body.Content);
                if (msg.Mentions.Any()) Console.WriteLine("Mentions: " + String.Join(", ", msg.Mentions.Select(m => m.Mentioned.User.DisplayName)));
                // Console.WriteLine("Summary: " + msg.Summary);
                Console.WriteLine("");
            }
            Console.WriteLine("");
        }

        static void GetChatId()
        {
            string chatIdInput = "";
            do
            {
                if (!String.IsNullOrWhiteSpace(chatIdInput))
                {
                    Console.WriteLine("Incorrect ChatId string format. Try again...");
                }
                Console.Write("Please enter chatId (Type 0 to cancel):");
                chatIdInput = Console.ReadLine();
            } while (!(chatIdInput == cancelString || Regex.IsMatch(chatIdInput, chatIdRegexPattern)));

            testChatId = chatIdInput;
        }
    }
}
