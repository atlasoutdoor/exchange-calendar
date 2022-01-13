Imports System
Imports System.Linq
Imports DayPilot.Web.Ui.Enums.Calendar
Imports DayPilot.Web.Ui.Events
Imports Microsoft.Exchange.WebServices.Data

Namespace TutorialVB
    Partial Public Class [Default]
        Inherits System.Web.UI.Page

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            LoadAppointments()
        End Sub

        Private ReadOnly Property Service() As ExchangeService
            Get
                Dim svc As New ExchangeService(ExchangeVersion.Exchange2010)
                svc.Credentials = New WebCredentials("user@yourcompany.com", "password")
                svc.Url = New Uri("https://outlook.office365.com/ews/exchange.asmx")
                Return svc
            End Get
        End Property

        Private Function FindDefaultCalendarFolder() As CalendarFolder
            Return CalendarFolder.Bind(Service, WellKnownFolderName.Calendar, New PropertySet())
        End Function


        Private Function FindNamedCalendarFolder(ByVal name As String) As CalendarFolder
            Dim view As New FolderView(100)
            view.PropertySet = New PropertySet(BasePropertySet.IdOnly)
            view.PropertySet.Add(FolderSchema.DisplayName)
            view.Traversal = FolderTraversal.Deep

            Dim sfSearchFilter As SearchFilter = New SearchFilter.IsEqualTo(FolderSchema.FolderClass, "IPF.Appointment")

            Dim findFolderResults As FindFoldersResults = Service.FindFolders(WellKnownFolderName.Root, sfSearchFilter, view)
            Return findFolderResults.Where(Function(f) f.DisplayName = name).Cast(Of CalendarFolder)().FirstOrDefault()
        End Function

        Protected Sub DayPilotCalendar1_OnTimeRangeSelected(ByVal sender As Object, ByVal e As TimeRangeSelectedEventArgs)
            Dim appointment As New Appointment(Service)

            appointment.Subject = "New Event"
            appointment.Start = e.Start
            appointment.End = e.End

            Dim folder As CalendarFolder = FindNamedCalendarFolder("Testing Calendar")

            appointment.Save(folder.Id, SendInvitationsMode.SendToNone)

            LoadAppointments()
        End Sub


        Private Sub LoadAppointments()
            Dim startDate As Date = DayPilot.Utils.Week.FirstDayOfWeek()
            Dim endDate As Date = startDate.AddDays(7)

            Dim calendar As CalendarFolder = FindNamedCalendarFolder("Testing Calendar") ' or FindDefaultCalendarFolder()

            Dim cView As New CalendarView(startDate, endDate, 50)
            cView.PropertySet = New PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id)
            Dim appointments As FindItemsResults(Of Appointment) = calendar.FindAppointments(cView)

            DayPilotCalendar1.ViewType = ViewTypeEnum.Week
            DayPilotCalendar1.DataStartField = "Start"
            DayPilotCalendar1.DataEndField = "End"
            DayPilotCalendar1.DataIdField = "Id"
            DayPilotCalendar1.DataTextField = "Subject"

            DayPilotCalendar1.DataSource = appointments
            DayPilotCalendar1.DataBind()

            DayPilotCalendar1.Update()

        End Sub

        Protected Sub DayPilotCalendar1_OnEventMove(ByVal sender As Object, ByVal e As EventMoveEventArgs)
            Dim appointment As Appointment = Microsoft.Exchange.WebServices.Data.Appointment.Bind(Service, New ItemId(e.Id))

            appointment.Start = e.NewStart
            appointment.End = e.NewEnd
            appointment.StartTimeZone = TimeZoneInfo.Local
            appointment.EndTimeZone = TimeZoneInfo.Local

            appointment.Update(ConflictResolutionMode.AlwaysOverwrite)

            LoadAppointments()
        End Sub
    End Class
End Namespace