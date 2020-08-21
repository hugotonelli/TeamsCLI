using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace TeamsCLI.Settings
{
    class ClientSettings
    {
        /// <summary>
        /// Display a notification for upcoming events when there is a reminder available
        /// </summary>
        public bool ReminderNotificationsEnabled { get; set; } = true;

        /// <summary>
        /// Time in minutes to snooze a reminder notififcation
        /// </summary>
        public int ReminderNotificationsSnoozeMinutes { get; set; } = 5;

        /// <summary>
        /// Add a timestamp to chat messages
        /// </summary>
        public bool ChatMessageTimestampEnabled { get; set; } = true;

        /// <summary>
        /// Format to use when calling ToString() on timestamp for chat messages from today
        /// </summary>
        public string ChatMessageTimestampFormatToday { get; set; } = "t";

        /// <summary>
        /// Format to use when calling ToString() on timestamp for chat messages from past days
        /// </summary>
        public string ChatMessageTimestampFormatPastDays { get; set; } = "g";
    }
}
