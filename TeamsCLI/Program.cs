using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Terminal.Gui;

namespace TeamsCLI
{
    class Program
    {
        const int maxChatsToDisplay = 25;
        static string chatIdRegexPattern = @"^\d+:[\d\w-_]+@(([\d\w-_]*)\.)+[\d\w-_]*$";

        private static Toplevel _top;
        private static MenuBar _menu;
        private static FrameView _leftPane;
        private static FrameView _rightPane;
        private static List<Chat> _chatList;
        private static ListView _chatsListView;
        private static int _currentSelectedChatIndex;
        private static string _currentChatId;
        private static Chat _currentChatInfo;
        private static List<ChatMessage> _chatMessageList;
        private static ListView _chatMessagesListView;
        private static Dialog _userSelectDialog;
        private static Terminal.Gui.Label _chatReplyLabel;
        private static ChatReplyTextField _chatReplyTextField;
        private static object _chatTimerToken;

        private static List<ConversationMember> _currentChatIdMembers;
        private static User _me;

        // Extends TextField to override ProcessColdKey to allow sending reply on Key.Enter
        private class ChatReplyTextField : TextField
        {
            public override bool ProcessColdKey(KeyEvent keyEvent)
            {
                if (this.HasFocus && keyEvent.Key == Key.Enter)
                {
                    Terminal.Gui.Application.MainLoop.Invoke(SendChatReply);
                    return true;
                }
                return base.ProcessColdKey(keyEvent);
            }
        }

        static void Main(string[] args)
        {

            Terminal.Gui.Application.Init();

            var myColorScheme = new ColorScheme()
            {
                Normal = Terminal.Gui.Attribute.Make(Color.White, Color.Black),
                Focus = Terminal.Gui.Attribute.Make(Color.Black, Color.Gray),
                HotNormal = Terminal.Gui.Attribute.Make(Color.Black, Color.DarkGray),
                HotFocus = Terminal.Gui.Attribute.Make(Color.White, Color.Black),
            };

            _top = Terminal.Gui.Application.Top;

            var statusBar = new StatusBar(new StatusItem[]
            {
                new StatusItem(Key.F1, "~F1~ Info", ShowInfo)
            });

            _top.Add(statusBar);

            _menu = new MenuBar(new MenuBarItem[]
            {
                new MenuBarItem("_File", new MenuItem[]
                {
                    new MenuItem("_Switch account", "", null),
                    new MenuItem("_Logout", "", null),
                    new MenuItem("_Quit", "", () => Terminal.Gui.Application.RequestStop()),
                }),
                new MenuBarItem("_Events", "", () => { }),
                new MenuBarItem("_Chats", new MenuItem[]
                {
                    new MenuItem("List chats", "", null),
                    new MenuItem("_New chat", "", null),
                })
            });

            _leftPane = new FrameView("Chats")
            {
                X = 0,
                Y = 1,
                Width = 25,
                Height = Dim.Fill(1),
                CanFocus = false,
                ColorScheme = myColorScheme,
            };

            _chatsListView = new ListView()
            {
                X = 0,
                Y = 0,
                Width = Dim.Fill(0),
                Height = Dim.Fill(0),
                AllowsMarking = false,
                CanFocus = true,
            };

            _chatsListView.OpenSelectedItem += (a) =>
            {
                _rightPane.SetFocus();
            };
            _chatsListView.SelectedItemChanged += (a) =>
            {
                _chatMessagesListView.Add(new View("testing"));
            };

            _leftPane.Add(_chatsListView);

            _rightPane = new FrameView("Messages")
            {
                X = 25,
                Y = 1,
                Width = Dim.Fill(),
                Height = Dim.Fill(1),
                CanFocus = true,
                ColorScheme = myColorScheme,
            };

            _chatMessagesListView = new ListView()
            {
                X = 0,
                Y = 0,
                Width = Dim.Fill(0),
                Height = Dim.Fill(1),
                AllowsMarking = false,
                CanFocus = true,
            };

            _chatReplyLabel = new Terminal.Gui.Label("Reply: ")
            {
                Y = Pos.Bottom(_chatMessagesListView),
                Width = 7
            };
            _chatReplyTextField = new ChatReplyTextField()
            {
                X = Pos.Right(_chatReplyLabel),
                Y = Pos.Bottom(_chatMessagesListView),
                Width = Dim.Fill(),
                CanFocus = true,
            };

            _rightPane.Add(_chatMessagesListView);
            _rightPane.Add(_chatReplyLabel);
            _rightPane.Add(_chatReplyTextField);

            _top.Add(_menu);
            _top.Add(_leftPane);
            _top.Add(_rightPane);

            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                var ok = new Button("Quit", true)
                {
                    Clicked = () => Terminal.Gui.Application.RequestStop()
                };
                var errorDialog = new Dialog("Error", ok);
                errorDialog.Text = "Missing or invalid appsettings.json";
                Terminal.Gui.Application.Run(errorDialog);
            }
            else
            {
                var appId = appConfig["appId"];
                var scopesString = appConfig["scopes"];
                var tenantId = appConfig["tenantId"];
                var scopes = scopesString.Split(';');

                // Initialize the auth provider with values from appsettings.json
                var authProvider = new DeviceCodeAuthProvider(appId, scopes, tenantId);

                var accounts = authProvider.GetAccounts().Result.ToArray();

                var cantAccounts = accounts.Count();
                if (cantAccounts > 1)
                {
                    Button ok;
                    ListView userListView;

                    var selectedIndex = -1;


                    while (selectedIndex == -1)
                    {
                        userListView = new ListView(accounts.Select(a => a.Username).ToList())
                        {
                            X = 0,
                            Y = 0,
                            Height = Dim.Fill(1),
                            Width = Dim.Fill(),
                            AllowsMarking = false,
                            CanFocus = true,
                            OpenSelectedItem = (a) =>
                            {
                                selectedIndex = a.Item;
                                Terminal.Gui.Application.RequestStop();
                            }
                        };

                        ok = new Button("OK", true)
                        {
                            Clicked = () =>
                            {
                                selectedIndex = userListView.SelectedItem;
                                Terminal.Gui.Application.RequestStop();
                            }
                        };
                        _userSelectDialog = new Dialog("Select account", ok)
                        {
                            Width = 40,
                            Height = 20,

                        };
                        _userSelectDialog.Add(userListView);
                        //_top.Add(_userSelectDialog);

                        Terminal.Gui.Application.Run(_userSelectDialog);
                    }

                    authProvider.SetAccount(accounts[selectedIndex]);
                }
                else
                {
                    authProvider.SetAccount(accounts.FirstOrDefault());
                }

                // Request a token to sign in the user
                var accessToken = authProvider.GetAccessToken().Result;

                // Initialize Graph client
                GraphHelper.Initialize(authProvider);

                Terminal.Gui.Application.MainLoop.Invoke(GetMe);

                //var chatList = GraphHelper.GetChatsAsync().ConfigureAwait(false);

                List<Chat> chats = new List<Chat>();


                Terminal.Gui.Application.MainLoop.Invoke(ShowChatList);

                Terminal.Gui.Application.Run(_top);
            }
        }

        private static async void GetMe()
        {
            var currentUser = await GraphHelper.GetMeAsync();
            _me = currentUser;
        }

        private static async void ShowChatList()
        {
            //_chatsListView.Source = null;
            var chats = await GraphHelper.GetChats();
            _chatList = chats.ToList();
            var ids = chats.Select(c => c.Id).ToList();
            _chatsListView.SetSource(ids);
            _chatsListView.OpenSelectedItem = (a) =>
            {
                _currentSelectedChatIndex = a.Item;
                _currentChatId = _chatList[_currentSelectedChatIndex].Id;
                Terminal.Gui.Application.MainLoop.Invoke(ShowChatMessages);
                Terminal.Gui.Application.MainLoop.Invoke(LoadChatInfo);
                Terminal.Gui.Application.MainLoop.Invoke(LoadChatMembers);
                //ShowChatMessages();
            };
        }

        private static async void ShowChatMessages()
        {
            var chatMessages = await GraphHelper.GetChatMessages(_currentChatId);
            _chatMessageList = chatMessages.ToList();
            RenderChatMessages();
            _chatMessagesListView.SetFocus();

            ResetChatMessagesTimer();
        }

        private static async void LoadChatInfo()
        {
            _currentChatInfo = null;
            var chatInfo = await GraphHelper.GetSingleChatAsync(_currentChatId);
            _currentChatInfo = chatInfo;
        }

        private static async void LoadChatMembers()
        {
            _currentChatIdMembers = null;
            var chatMembers = await GraphHelper.GetChatMembers(_currentChatId);
            _currentChatIdMembers = chatMembers.ToList();
        }

        private static void ResetChatMessagesTimer()
        {
            if (_chatTimerToken != null)
            {
                Terminal.Gui.Application.MainLoop.RemoveTimeout(_chatTimerToken);
            }

            bool timer(MainLoop caller)
            {
                RefreshChatMessages();
                return true;
            }
            _chatTimerToken = Terminal.Gui.Application.MainLoop.AddTimeout(TimeSpan.FromSeconds(20), timer);
        }

        private static void RenderChatMessages()
        {
            int msgCount;
            var msgs = _chatMessageList.Select(m =>
                {
                    string modifiers = "";

                    if (m.Mentions != null && m.Mentions.Any(n => n.Mentioned?.User.Id == _me.Id))
                    {
                        modifiers = String.Concat(modifiers, "@");
                    }

                    if (m.Importance.HasValue && m.Importance.Value != ChatMessageImportance.Normal)
                    {
                        modifiers = String.Concat(modifiers, "!");
                    }

                    if (!String.IsNullOrEmpty(modifiers))
                    {
                        modifiers = String.Concat(modifiers, " ");
                    }

                    return modifiers
                        + $"[{m.CreatedDateTime?.ToLocalTime().ToString("g")}] "
                        + m.From.User.DisplayName
                        + ": " + m.Body.Content;
                }).Reverse().ToList();

            if (msgs.Any())
            {
                _chatMessagesListView.SetSource(msgs);
                msgCount = msgs.Count;
            }
            else
            {
                var noMessagesList = new List<string>() { "==== No messages to display ====" };
                msgCount = 1;
                _chatMessagesListView.SetSource(noMessagesList);
            }


            var topItem = msgCount < _chatMessagesListView.Frame.Height
                ? 0
                : msgCount - _chatMessagesListView.Frame.Height;
            _chatMessagesListView.SelectedItem = msgCount - 1;
            _chatMessagesListView.TopItem = topItem;
        }

        private static async void RefreshChatMessages()
        {
            var lastMessageDate = _chatMessageList.FirstOrDefault()?.CreatedDateTime;
            var chatMessages = await GraphHelper.GetChatMessages(_currentChatId);

            if (lastMessageDate != null && lastMessageDate.HasValue)
            {
                var newMessages = chatMessages.Where(m => m.CreatedDateTime > lastMessageDate);
                _chatMessageList.InsertRange(0, newMessages);
            }
            else
            {
                _chatMessageList = chatMessages.ToList();
            }
            RenderChatMessages();
        }

        private static async void SendChatReply()
        {
            if (!_chatReplyTextField.Text.IsEmpty)
            {
                await GraphHelper.PostChatMessage(_currentChatId, _chatReplyTextField.Text.ToString());
                _chatReplyTextField.Text = String.Empty;

                ResetChatMessagesTimer();
                Terminal.Gui.Application.MainLoop.AddTimeout(TimeSpan.FromMilliseconds(500), (mainloop) =>
                {
                    RefreshChatMessages();
                    return false;
                });
            }
            Console.WriteLine("Invoked Chat Reply method!");
        }

        private static void ShowInfo()
        {
            if (String.IsNullOrEmpty(_currentChatId) || _currentChatIdMembers == null)
            {
                return;
            }

            var ok = new Button("OK", true)
            {
                Clicked = () => Terminal.Gui.Application.RequestStop()
            };

            string text = "\n ChatId: " + _currentChatId
                    + "\n Topic: " + (String.IsNullOrEmpty(_currentChatInfo?.Topic) ? "[No topic]" : _currentChatInfo.Topic)
                    + "\n\n Members:\n" + String.Join(", ", _currentChatIdMembers.Select(m => m.DisplayName));

            var infoDialog = new Dialog("Chat Info", ok)
            {
                ColorScheme = Colors.Menu,
                Text = text,
            };

            Terminal.Gui.Application.Run(infoDialog);

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
    }
}
