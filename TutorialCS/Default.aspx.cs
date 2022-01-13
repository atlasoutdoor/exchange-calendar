using System;
using System.Linq;
using DayPilot.Web.Ui.Enums.Calendar;
using DayPilot.Web.Ui.Events;
using Microsoft.Exchange.WebServices.Data;

namespace TutorialCS
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            LoadAppointments();
        }

        private ExchangeService Service
        {
            get
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.Credentials = new WebCredentials("user@yourcompany.com", "password");
                service.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");
                return service;
            }
        }

        private CalendarFolder FindDefaultCalendarFolder()
        {
            return CalendarFolder.Bind(Service, WellKnownFolderName.Calendar, new PropertySet());
        }


        private CalendarFolder FindNamedCalendarFolder(string name)
        {
            FolderView view = new FolderView(100);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;

            SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(FolderSchema.FolderClass, "IPF.Appointment");

            FindFoldersResults findFolderResults = Service.FindFolders(WellKnownFolderName.Root, sfSearchFilter, view);
            return findFolderResults.Where(f => f.DisplayName == name).Cast<CalendarFolder>().FirstOrDefault();
        }

        protected void DayPilotCalendar1_OnTimeRangeSelected(object sender, TimeRangeSelectedEventArgs e)
        {
            Appointment appointment = new Appointment(Service);

            appointment.Subject = "New Event";
            appointment.Start = e.Start;
            appointment.End = e.End;

            CalendarFolder folder = FindNamedCalendarFolder("Testing Calendar");

            appointment.Save(folder.Id, SendInvitationsMode.SendToNone);

            LoadAppointments();
        }


        private void LoadAppointments()
        {
            DateTime startDate = DayPilot.Utils.Week.FirstDayOfWeek();
            DateTime endDate = startDate.AddDays(7);

            CalendarFolder calendar = FindNamedCalendarFolder("Testing Calendar");  // or FindDefaultCalendarFolder()

            CalendarView cView = new CalendarView(startDate, endDate, 50);
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            DayPilotCalendar1.ViewType = ViewTypeEnum.Week;
            DayPilotCalendar1.DataStartField = "Start";
            DayPilotCalendar1.DataEndField = "End";
            DayPilotCalendar1.DataIdField = "Id";
            DayPilotCalendar1.DataTextField = "Subject";

            DayPilotCalendar1.DataSource = appointments;
            DayPilotCalendar1.DataBind();

            DayPilotCalendar1.Update();

        }

        protected void DayPilotCalendar1_OnEventMove(object sender, EventMoveEventArgs e)
        {
            Appointment appointment = Appointment.Bind(Service, new ItemId(e.Id));

            appointment.Start = e.NewStart;
            appointment.End = e.NewEnd;
            appointment.StartTimeZone = TimeZoneInfo.Local;
            appointment.EndTimeZone = TimeZoneInfo.Local;

            appointment.Update(ConflictResolutionMode.AlwaysOverwrite);

            LoadAppointments();
        }
    }
}