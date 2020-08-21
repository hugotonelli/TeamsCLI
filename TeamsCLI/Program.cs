using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Terminal.Gui;
using TeamsCLI.Settings;
using NStack;

namespace TeamsCLI
{
    class Program
    {
        const int maxChatsToDisplay = 25;
        static string chatIdRegexPattern = @"^\d+:[\d\w-_]+@(([\d\w-_]*)\.)+[\d\w-_]*$";

        private static DeviceCodeAuthProvider _authProvider;
        private static string _accessToken;

        private static Dialog _authTokenDialog;

        private static Toplevel _top;
        private static MenuBar _menu;
        private static FrameView _leftPane;
        private static FrameView _rightPane;
        private static List<Chat> _chatList;
        private static List<string> _chatListStrings;
        private static List<string> _chatIdsWithNoTopicOrMembers;
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

        private static Dialog _eventsCalendarDialog;
        private static List<Event> _eventList = new List<Event>();
        private static List<Reminder> _eventReminderList = new List<Reminder>();
        private static ScrollView _eventsScrollView;
        private static List<Meeting> _meetingsList = new List<Meeting>();

        private static ColorScheme _colorScheme;
        private static ColorScheme _colorSchemeInverse;

        private static List<ConversationMember> _currentChatIdMembers;
        private static User _me;

        private static ClientSettings settings;


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

        public static Action running = MainApp;
        static void Main(string[] args)
        {
            while (running != null)
            {
                running.Invoke();
            }
            Terminal.Gui.Application.Shutdown();
        }

        static void MainApp()
        {
            settings = SettingsManager.ReadOrCreateSettings();

            Terminal.Gui.Application.Init();

            // Builds the color schemes for the windows after Gui app init
            _colorScheme = new ColorScheme()
            {
                Normal = Terminal.Gui.Attribute.Make(Color.White, Color.Black),
                Focus = Terminal.Gui.Attribute.Make(Color.White, Color.Blue),
                HotNormal = Terminal.Gui.Attribute.Make(Color.Black, Color.Gray),
                HotFocus = Terminal.Gui.Attribute.Make(Color.White, Color.Black),
            };
             _colorSchemeInverse = new ColorScheme()
             {
                 Normal = Terminal.Gui.Attribute.Make(Color.White, Color.Black),
                 Focus = Terminal.Gui.Attribute.Make(Color.Black, Color.Gray),
                 HotNormal = Terminal.Gui.Attribute.Make(Color.White, Color.Blue),
                 HotFocus = Terminal.Gui.Attribute.Make(Color.White, Color.Black),
             };

            _top = Terminal.Gui.Application.Top;

            #region Setup

            var appConfig = LoadAppConfig();

            // Shows error and terminates app if appConfig is null
            if (appConfig == null)
            {
                MessageBox.ErrorQuery(40, 5, "Error", "Missing or invalid appsettings.json", "OK");
                running = null;
                _top.Running = false;
                return;
            }

            #endregion Setup

            #region Login

            // Gets settings required for DeviceCodeAuthProvider
            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var tenantId = appConfig["tenantId"];
            var scopes = scopesString.Split(';');

            // Initialize the auth provider
            _authProvider = new DeviceCodeAuthProvider(
                appId, scopes, tenantId,
                (callBack) => {
                    // Display the DeviceCodeResult message in a MessageBox
                    MessageBox.Query(50, 8, 
                        "Remote Device Sign-in",
                        callBack.Message + " Once you have successfully signed in, hit OK.",
                        "OK");

                    return Task.FromResult(0);
                });

            var accounts = _authProvider.GetAccounts().Result.ToArray();

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
                            running = null;
                            Terminal.Gui.Application.RequestStop();
                        }
                    };

                    ok = new Button("OK", true)
                    {
                        Clicked = () =>
                        {
                            selectedIndex = userListView.SelectedItem;
                            running = null;
                            Terminal.Gui.Application.RequestStop();
                        }
                    };
                    _userSelectDialog = new Dialog("Select account", ok)
                    {
                        Width = 40,
                        Height = 10,

                    };
                    _userSelectDialog.Add(userListView);
                    //_top.Add(_userSelectDialog);

                    Terminal.Gui.Application.Run(_userSelectDialog);
                }

                _authProvider.SetAccount(accounts[selectedIndex]);
            }
            else
            {
                _authProvider.SetAccount(accounts.FirstOrDefault());
            }

            // Request a token to sign in the user
            GetAccessToken();

            // Initialize Graph client
            GraphHelper.Initialize(_authProvider);

            Terminal.Gui.Application.MainLoop.Invoke(GetMe);



            #endregion Login

            #region Main Window

            static void Quit()
            {
                running = null;
                _top.Running = false;
                Terminal.Gui.Application.RequestStop();
            };

            var statusBar = new StatusBar(new StatusItem[]
            {
                new StatusItem(Key.F1, "~F1~ Info", ShowInfo),
                new StatusItem(Key.F5, "~F5~ Events", ShowEvents),
                new StatusItem(Key.ControlQ, "~^Q~ Quit", Quit),
            });

            _top.Add(statusBar);

            _menu = new MenuBar(new MenuBarItem[]
            {
                new MenuBarItem("_File", new MenuItem[]
                {
                    //new MenuItem("_Switch account", "", null),
                    //new MenuItem("_Logout", "", null),
                    new MenuItem("_Quit", "", Quit),
                }),
                //new MenuBarItem("_Events", "", () => { }),
                //new MenuBarItem("_Chats", new MenuItem[]
                //{
                //    new MenuItem("List chats", "", null),
                //    new MenuItem("_New chat", "", null),
                //}),
                new MenuBarItem("_Settings", "", ShowSettings),
            });

            _leftPane = new FrameView("Chats")
            {
                X = 0,
                Y = 1,
                Width = Dim.Percent(20),
                Height = Dim.Fill(1),
                CanFocus = false,
                ColorScheme = _colorScheme,
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

            _leftPane.Add(_chatsListView);

            _rightPane = new FrameView("Messages")
            {
                X = Pos.Right(_leftPane),
                Y = 1,
                Width = Dim.Fill(),
                Height = Dim.Fill(1),
                CanFocus = true,
                ColorScheme = _colorScheme,
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
                ColorScheme = Colors.Base,
            };

            _rightPane.Add(_chatMessagesListView);
            _rightPane.Add(_chatReplyLabel);
            _rightPane.Add(_chatReplyTextField);

            _top.Add(_menu);
            _top.Add(_leftPane);
            _top.Add(_rightPane);

            #endregion Main Window

            List<Chat> chats = new List<Chat>();

            Terminal.Gui.Application.MainLoop.Invoke(ShowChatList);

            RegisterReminderTimer();

            Terminal.Gui.Application.Run(_top);
        }

        private static void GetAccessToken()
        {
            while (String.IsNullOrEmpty(_accessToken))
            {
                var accessToken = _authProvider.GetAccessTokenBlocking();
                _accessToken = accessToken;
            }
        }

        private static async void GetMe()
        {
            var currentUser = await GraphHelper.GetMeAsync();
            _me = currentUser;
        }

        private static object RegisterReminderTimer()
        {
            if (!settings.ReminderNotificationsEnabled)
            {
                return null;
            }

            bool reminderTimer(MainLoop caller)
            {
                CheckReminders();
                return true;
            }

            return Terminal.Gui.Application.MainLoop.AddTimeout(TimeSpan.FromMinutes(1), reminderTimer);
        }


        private static async void CheckReminders()
        {
            await UpdateMeetingsList();

            DateTime now = DateTime.UtcNow;
            int snoozeMins = settings.ReminderNotificationsSnoozeMinutes;
            var reminders = _meetingsList
                .Where(m =>
                    !m.Dismissed
                    && m.Reminder.ReminderFireTime.ToDateTime() <= now
                    && m.Event.Start.ToDateTime().AddMinutes(15) >= now 
                    && (m.SnoozedUntil == null
                        || (m.SnoozedUntil.HasValue && m.SnoozedUntil.Value <= now)));
            if (reminders.Any())
            {
                var msg = reminders.Aggregate("", (aggr, curr) =>
                    aggr + $"{(curr.Event.Start.ToDateTime() - now).TotalMinutes:#} mins : {curr.Event.Subject}\n");
                var action = MessageBox.Query("Reminders", msg, $"Snooze {snoozeMins} mins", "Dismiss All"
                );

                if (action == 1)
                {
                    foreach(var reminder in reminders)
                    {
                        reminder.Dismissed = true;
                    }
                }
                else
                {
                    var snoozeUntil = now.AddMinutes(snoozeMins);
                    foreach(var reminder in reminders)
                    {
                        reminder.SnoozedUntil = snoozeUntil;
                    }
                }
            }
        }

        private static async void ShowChatList()
        {
            //_chatsListView.Source = null;
            var chats = await GraphHelper.GetChats();
            _chatList = chats.ToList();
            _chatListStrings = chats.Select(c =>
                {
                    if (c.Topic != null && !String.IsNullOrEmpty(c.Topic)) return c.Topic;
                    if (c.Members != null && c.Members.Any()) return String.Join(", ", c.Members.Select(m => m.DisplayName).Where(n => n != _me.DisplayName));
                    return c.Id;
                }).ToList();

            _chatsListView.SetSource(_chatListStrings);

            _chatIdsWithNoTopicOrMembers = _chatList.FindAll(c => c.Topic == null && c.Members == null).Select(c => c.Id).ToList();
            Terminal.Gui.Application.MainLoop.Invoke(RefreshMembersInChatList);

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

        private static async void RefreshMembersInChatList()
        {
            string myDisplayName = _me.DisplayName;
            foreach(string chatId in _chatIdsWithNoTopicOrMembers)
            {
                var members = await GraphHelper.GetChatMembers(chatId);
                var membersString = (members == null) ? String.Empty : String.Join(", ", members.Select(m => m.DisplayName).Where(m => m != myDisplayName));
                if (!String.IsNullOrEmpty(membersString))
                {
                    int index = _chatListStrings.IndexOf(chatId);
                    _chatListStrings[index] = membersString;
                }
            }
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
            _currentChatIdMembers = chatMembers?.ToList();
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
                    string modifiers = String.Empty;
                    string timeTag = String.Empty;
                    if (settings.ChatMessageTimestampEnabled)
                    {
                        string timeStamp = (m.CreatedDateTime?.ToLocalTime().Date == DateTime.Today)?
                                m.CreatedDateTime?.ToLocalTime().ToString(settings.ChatMessageTimestampFormatToday)
                                : m.CreatedDateTime?.ToLocalTime().ToString(settings.ChatMessageTimestampFormatPastDays);
                        timeTag = $"[{timeStamp}] ";
                    }
                        

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
                        + timeTag
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

        private static void ShowSettings()
        {
            var reminderNotificationsEnabledLabel = new Terminal.Gui.Label("Display reminder notifications ");
            var reminderNotificationsEnabledCheckBox = new CheckBox(String.Empty, settings.ReminderNotificationsEnabled)
            {
                X = Pos.Right(reminderNotificationsEnabledLabel),
                //Y = Pos.Bottom(reminderNotificationsLabel) + 1,
            };
            var reminderNotificationsSnoozeMinutesLabel = new Terminal.Gui.Label("Minutes to snooze: ")
            {
                Y = Pos.Bottom(reminderNotificationsEnabledCheckBox) + 1,
            };
            var reminderNotificationsSnoozeMinutesTextField = new TextField($"{settings.ReminderNotificationsSnoozeMinutes}")
            {
                X = Pos.Right(reminderNotificationsSnoozeMinutesLabel),
                Y = Pos.Bottom(reminderNotificationsEnabledCheckBox) + 1,
                Width = 4,
                ColorScheme = _colorSchemeInverse,
            };
            reminderNotificationsSnoozeMinutesTextField.TextChanged += (args) =>
            {
                var text = reminderNotificationsSnoozeMinutesTextField.Text.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    text = new Regex(@"[^\d]+").Replace(text, "");
                    reminderNotificationsSnoozeMinutesTextField.Text = text;
                }
            };
            var reminderNotificationsFrameView = new FrameView("Reminder notifications")
            {
                X = 0,
                Y = 0,
                Height = 5,
                Width = Dim.Fill(),
            };
            reminderNotificationsFrameView.Add(
                reminderNotificationsEnabledLabel,
                reminderNotificationsEnabledCheckBox,
                reminderNotificationsSnoozeMinutesLabel,
                reminderNotificationsSnoozeMinutesTextField
                );

            var chatMessageTimestampEnabledLabel = new Terminal.Gui.Label("Display timestamps in chat messages");
            var chatMessageTimestampEnabledCheckBox = new CheckBox(String.Empty, settings.ChatMessageTimestampEnabled)
            {
                X = Pos.Right(chatMessageTimestampEnabledLabel),
                //Y = Pos.Bottom(chatMessageLabel) + 1,
            };
            var timeFormatRadioItems = new ustring[] { "Date & Time", "Date only", "Time only" };
            var timeFormatStringRadioItems = new List<string> { "g", "d", "t" };
            var chatMessageTimestampFormatTodayLabel = new Terminal.Gui.Label("Timestamp format (today): ")
            {
                Y = Pos.Bottom(chatMessageTimestampEnabledCheckBox) + 1,
            };
            var chatMessageTimestampFormatTodayRadioGroup = new RadioGroup(timeFormatRadioItems)
            {
                X = Pos.Right(chatMessageTimestampFormatTodayLabel),
                Y = Pos.Bottom(chatMessageTimestampEnabledCheckBox) + 1,
                SelectedItem = timeFormatStringRadioItems.IndexOf(settings.ChatMessageTimestampFormatToday),
            };
            var chatMessageTimestampFormatPastDaysLabel = new Terminal.Gui.Label("Timestamp format (past): ")
            {
                Y = Pos.Bottom(chatMessageTimestampFormatTodayRadioGroup) + 1,
            };
            var chatMessageTimestampFormatPastDaysRadioGroup = new RadioGroup(timeFormatRadioItems)
            {
                X = Pos.Right(chatMessageTimestampFormatPastDaysLabel),
                Y = Pos.Bottom(chatMessageTimestampFormatTodayRadioGroup) + 1,
                SelectedItem = timeFormatStringRadioItems.IndexOf(settings.ChatMessageTimestampFormatPastDays),
            };
            var chatMessageFrameView = new FrameView("Chat messages")
            {
                Y = Pos.Bottom(reminderNotificationsFrameView),
                Height = 11,
                Width = Dim.Fill(),
            };
            chatMessageFrameView.Add(
                chatMessageTimestampEnabledLabel,
                chatMessageTimestampEnabledCheckBox,
                chatMessageTimestampFormatTodayLabel,
                chatMessageTimestampFormatTodayRadioGroup,
                chatMessageTimestampFormatPastDaysLabel,
                chatMessageTimestampFormatPastDaysRadioGroup
                );

            var saveButton = new Button("Save", false)
            {
                Clicked = () =>
                {
                    settings.ReminderNotificationsEnabled = reminderNotificationsEnabledCheckBox.Checked;

                    if (int.TryParse(reminderNotificationsSnoozeMinutesTextField.Text.ToString(), out int mins))
                    {
                        settings.ReminderNotificationsSnoozeMinutes = mins;
                    }

                    settings.ChatMessageTimestampEnabled = chatMessageTimestampEnabledCheckBox.Checked;

                    settings.ChatMessageTimestampFormatToday = 
                        timeFormatStringRadioItems[chatMessageTimestampFormatTodayRadioGroup.SelectedItem];

                    settings.ChatMessageTimestampFormatPastDays =
                        timeFormatStringRadioItems[chatMessageTimestampFormatPastDaysRadioGroup.SelectedItem];

                    SettingsManager.SaveSettings(settings);
 
                    Terminal.Gui.Application.RequestStop();
                }
            };
            var cancelButton = new Button("Cancel", false)
            {
                Clicked = () =>
                {
                    Terminal.Gui.Application.RequestStop();
                }
            };
            var settingsDialog = new Dialog("Settings", saveButton, cancelButton)
            {
                ColorScheme = _colorScheme,
                Width = 75,
            };

            settingsDialog.Add(
                reminderNotificationsFrameView,
                //reminderNotificationsLabel,
                //reminderNotificationsEnabledCheckBox,
                //reminderNotificationsSnoozeMinutesLabel,
                //reminderNotificationsSnoozeMinutesTextField,
                chatMessageFrameView
                //chatMessageLabel,
                //chatMessageTimestampEnabledCheckBox,
                //chatMessageTimestampFormatTodayLabel,
                //chatMessageTimestampFormatTodayRadioGroup,
                //chatMessageTimestampFormatPastDaysLabel,
                //chatMessageTimestampFormatPastDaysRadioGroup
                );

            Terminal.Gui.Application.Run(settingsDialog);
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

        private static void ShowEvents()
        {
            //Terminal.Gui.Application.MainLoop.Invoke(ListCalendarEvents);
            Terminal.Gui.Application.MainLoop.Invoke(ListEventsWithReminderEvents);
            var close = new Button("Close")
            {
                Clicked = () => Terminal.Gui.Application.RequestStop(),
            };

            _eventsCalendarDialog = new Dialog("Upcoming Events", close)
            {
                Y = Pos.Center(),
                Width = 63,
                ColorScheme = Colors.Menu,
            };

            _eventsScrollView = new ScrollView()
            {
                X = 0,
                Y = 0,
                ColorScheme = _colorScheme,
                Width = Dim.Fill(),
                Height = Dim.Fill(1),
                ShowVerticalScrollIndicator = true,
                //ShowHorizontalScrollIndicator = true,
                AutoHideScrollBars = false,
            };

            _eventsCalendarDialog.Add(_eventsScrollView);

            var textLabel = new Terminal.Gui.Label("No events available.");
            _eventsScrollView.ContentSize = new Size(61, _eventList.Count);

            _eventsScrollView.Add(textLabel);

            Terminal.Gui.Application.Run(_eventsCalendarDialog);
        }

        static IConfigurationRoot LoadAppConfig()
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

        private static async void ListCalendarEvents()
        {
            _eventList = new List<Event>();
            //var events = await GraphHelper.GetEventsAsync();
            var events = await GraphHelper.GetCalendarItemsAsync();
            //var events = await GraphHelper.GetEventsInAllCalendarsAsync();
            //var reminders = await GraphHelper.GetRemindersAsync();

            if (events == null)
            {
                return;
            }

            _eventList = events.ToList();
            _eventsScrollView.RemoveAll();

            for (int i = 0; i < _eventList.Count; i++)
            {
                var eventItem = _eventList[i];
                var newEventWindow = new FrameView(eventItem.Subject)
                {
                    X = 0,
                    Y = 6 * i,
                    Height = 6,
                    Width = Dim.Fill(),
                    ColorScheme = _colorScheme,
                    CanFocus = true,
                    TextAlignment = TextAlignment.Left,
                    Text = $"Organizer: {eventItem.Organizer.EmailAddress.Name}\n"
                            + $"Start: {FormatDateTimeTimeZone(eventItem.Start)}\n"
                            + $"End: {FormatDateTimeTimeZone(eventItem.End)}\n"
                };
                var detailsButton = new Button("Details", true)
                {
                    Y = 3,
                    X = Pos.Center(),
                    Clicked = () =>
                    {
                        string attendees;
                        if (eventItem.Attendees == null)
                        {
                            attendees = "No details available.";
                        }
                        else
                        {
                            attendees = "Attendees: "
                                + String.Join(", ", eventItem.Attendees?.Select(a => a.EmailAddress.Name));
                        }
                        MessageBox.Query(eventItem.Subject, attendees, "OK");
                    }
                };
                newEventWindow.Enter += (e) => 
                    {
                        int absY = Math.Abs(_eventsScrollView.ContentOffset.Y);

                        if (absY > newEventWindow.Frame.Y)
                        {
                            _eventsScrollView.ContentOffset = new Point(0, -newEventWindow.Frame.Y);
                        }
                        else if (absY + _eventsScrollView.Frame.Height < newEventWindow.Frame.Bottom)
                        {
                            _eventsScrollView.ContentOffset = new Point(0,
                                -(newEventWindow.Frame.Bottom - _eventsScrollView.Frame.Height));
                        }
                    };
                newEventWindow.Add(detailsButton);
                _eventsScrollView.Add(newEventWindow);
            }

            _eventsScrollView.ContentSize = new Size(60, _eventList.Count * 6);
        }

        private static async Task UpdateMeetingsList()
        {
            var meetings = await GraphHelper.GetEventsWithRemindersAsync();

            if (meetings == null)
            {
                return;
            }

            var meetingsDict = _meetingsList.ToDictionary(m => m.Event.Id);
            foreach (var meeting in meetings)
            {
                var id = meeting.Event.Id;
                if (meetingsDict.TryGetValue(id, out var meet))
                {
                    meetingsDict[id].Event = meet.Event;
                    meetingsDict[id].Reminder = meet.Reminder;
                }
                else
                {
                    meetingsDict[id] = meeting;
                }
            }
            _meetingsList = meetingsDict.Values.ToList();
        }

        private static async void ListEventsWithReminderEvents()
        {
            await UpdateMeetingsList();

            _eventsScrollView.RemoveAll();

            for (int i = 0; i < _meetingsList.Count; i++)
            {
                var meetingItem = _meetingsList[i];
                var newEventWindow = new FrameView(meetingItem.Event.Subject)
                {
                    X = 0,
                    Y = 6 * i,
                    Height = 6,
                    Width = Dim.Fill(),
                    ColorScheme = _colorSchemeInverse,
                    CanFocus = true,
                    TextAlignment = TextAlignment.Left,
                    Text = $"Organizer: {meetingItem.Event.Organizer.EmailAddress.Name}\n"
                            + $"Start: {FormatDateTimeTimeZone(meetingItem.Event.Start)}\n"
                            + $"End: {FormatDateTimeTimeZone(meetingItem.Event.End)}\n"
                };
                var detailsButton = new Button("Details", true)
                {
                    Y = 3,
                    X = Pos.Center(),
                    ColorScheme = _colorSchemeInverse,
                    Clicked = () =>
                    {
                        string details = "";

                        if (meetingItem.Event.Attendees == null)
                        {
                            details += "No attendees details available.\n";
                        }
                        else
                        {
                            details += "Attendees: "
                                + String.Join(", ", meetingItem.Event.Attendees?.Select(a => a.EmailAddress.Name))
                                + "\n";
                        }

                        details += "Reminder at: " + FormatDateTimeTimeZone(meetingItem.Reminder.ReminderFireTime);

                        MessageBox.Query(meetingItem.Event.Subject, details, "OK");
                    }
                };
                newEventWindow.Enter += (e) =>
                {
                    int absY = Math.Abs(_eventsScrollView.ContentOffset.Y);

                    if (absY > newEventWindow.Frame.Y)
                    {
                        _eventsScrollView.ContentOffset = new Point(0, -newEventWindow.Frame.Y);
                    }
                    else if (absY + _eventsScrollView.Frame.Height < newEventWindow.Frame.Bottom)
                    {
                        _eventsScrollView.ContentOffset = new Point(0,
                            -(newEventWindow.Frame.Bottom - _eventsScrollView.Frame.Height));
                    }
                };
                newEventWindow.Add(detailsButton);
                _eventsScrollView.Add(newEventWindow);
            }

            _eventsScrollView.ContentSize = new Size(60, _meetingsList.Count * 6);
        }
    }
}
