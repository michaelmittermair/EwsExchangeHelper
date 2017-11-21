using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace EwsExchangeHelper
{
    public partial class MsExchangeServices
    {
        /// <summary>
        /// Getting an appointment based on the id
        /// </summary>
        /// <param name="appointmentId">id of the appointment</param>
        /// <returns>appointment with the id</returns>
        public Appointment GetAppointmentById(ItemId appointmentId)
        {
            // Instantiate an appointment object by binding to it by using the ItemId.
            // As a best practice, limit the properties returned to only the ones you need.
            return Appointment.Bind(ExchangeService, appointmentId,
                new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End));

        }

        /// <summary>
        /// Adding an required attendee to the appointment
        /// </summary>
        /// <param name="mailAddress">e-mail-address of the user</param>
        /// <param name="appointmenId">id of the appointment</param>
        /// <param name="sendInvitationsMode">invitationmode</param>
        public void AddRequiredAttendee(string mailAddress, ItemId appointmenId,
            SendInvitationsMode sendInvitationsMode = SendInvitationsMode.SendToNone)
        {
            var appointment = GetAppointmentById(appointmenId);

            if (appointment != null)
            {
                appointment.RequiredAttendees.Add(mailAddress);
                appointment.Save(sendInvitationsMode);
            }
            else
            {
                throw new AppointmentNotFoundException(appointmenId);
            }
        }

        /// <summary>
        /// Retrieving all Calender events, maximum 1000 Elements.
        /// Filtering by start and enddate, all events between these dates will be returned
        /// </summary>
        /// <param name="startDateTime">startdate</param>
        /// <param name="endDateTime">enddate</param>
        /// <param name="numberOfAppointments">number of events</param>
        /// <returns>list of events</returns>
        public IEnumerable<Appointment> GetCalendarEvents(DateTime startDateTime, DateTime endDateTime,
            int numberOfAppointments = 0)
        {
            // Initialize the calendar folder object with only the folder ID. 
            var calendar = CalendarFolder.Bind(ExchangeService, WellKnownFolderName.Calendar, new PropertySet());

            // Set the start and end time and number of appointments to retrieve.
            var cView = new CalendarView(startDateTime, endDateTime)
            {
                // Limit the properties returned to the appointment's subject, start time, and end time.
                PropertySet = new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End)
            };

            if (numberOfAppointments > 0)
                cView.MaxItemsReturned = numberOfAppointments;

            // Retrieve a collection of appointments by using the calendar view.
            return calendar.FindAppointments(cView);
        }

        /// <summary>
        /// Creating an appointment
        /// </summary>
        /// <param name="subject">Subject of the appointment</param>
        /// <param name="body">Text/Description of the appointment</param>
        /// <param name="start">Startdate and time</param>
        /// <param name="end">Enddate and time</param>
        /// <param name="invitationsMode">SendInvitationsMode</param>
        public void CreateCalenderEvent(string subject, string body, DateTime start, DateTime end,
            SendInvitationsMode invitationsMode)
        {
            var appointment = new Appointment(ExchangeService)
            {
                Start = start,
                End = end,
                Subject = subject,
                Body = body,
                Location = body
            };
            appointment.Save(invitationsMode);
        }
    }
}
