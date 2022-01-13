<%@ Page Title="" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="TutorialVB.Default" %>
<%@ Register Assembly="DayPilot" Namespace="DayPilot.Web.Ui" TagPrefix="DayPilot" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server"></asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
		<div>
			<DayPilot:DayPilotCalendar runat="server" id="DayPilotCalendar1"

				TimeRangeSelectedHandling="CallBack"
				OnTimeRangeSelected="DayPilotCalendar1_OnTimeRangeSelected"

				EventMoveHandling = "CallBack"
				OnEventMove="DayPilotCalendar1_OnEventMove"
				>                
			</DayPilot:DayPilotCalendar>

		</div>    
</asp:Content>