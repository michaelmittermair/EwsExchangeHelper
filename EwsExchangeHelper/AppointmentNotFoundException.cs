using System;
using Microsoft.Exchange.WebServices.Data;

namespace EwsExchangeHelper
{
    public class AppointmentNotFoundException : Exception
    {
        public ItemId AppointmentId { get; set; }

        public AppointmentNotFoundException(ItemId appointmentId)
        {
            AppointmentId = appointmentId;
        }
    }
}
