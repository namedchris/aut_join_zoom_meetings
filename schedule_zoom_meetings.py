import datetime
from datetime import timedelta, timezone
import win32com.client
import webbrowser

def get_calendar(begin,end):
	outlook = win32com.client.Dispatch('Outlook.Application').getNamespace('MAPI')
	calendar = outlook.getDefaultFolder(9).Items
	calendar.IncludeRecurrences = True
	calendar.Sort('[Start]')
	restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
	calendar = calendar.Restrict(restriction)
	return calendar

def get_scheduler():
	scheduler = win32com.client.Dispatch('Schedule.Service')
	scheduler.Connect()
	return scheduler

def schedule_zoom_meetings(calendar,scheduler):
	root_folder = scheduler.GetFolder('zoom_meetings')
	for appointment in calendar:
		if "zoom.us" in appointment.location:
			task_def = scheduler.NewTask(0)
			# I don't want to mess with pytz right now.
			# I think win32.com may have functionality to do this as well
			# Also, task scheduler has a setting to turn time zones on
			start_time = appointment.start + timedelta(hours=6)
			# Create trigger
			TASK_TRIGGER_TIME = 1
			trigger = task_def.Triggers.Create(TASK_TRIGGER_TIME)
			trigger.StartBoundary = start_time.isoformat()
			# Create action
			TASK_ACTION_EXEC = 0
			action = task_def.Actions.Create(TASK_ACTION_EXEC)
			action.ID = appointment.subject + str(appointment.start)
			action.Path = 'python'
			action.Arguments = '-m webbrowser -t ' + appointment.location
			# Set parameters
			task_def.RegistrationInfo.Description = appointment.subject
			task_def.Settings.Enabled = True
			task_def.Settings.StopIfGoingOnBatteries = False

			# Register task
			# If task already exists, it will be updated
			TASK_CREATE_OR_UPDATE = 6
			TASK_LOGON_NONE = 0
			root_folder.RegisterTaskDefinition(
			    appointment.subject+str(hash(appointment.start)),  # Task name
			    task_def,
			    TASK_CREATE_OR_UPDATE,
			    '',  # No user
			    '',  # No password
			    TASK_LOGON_NONE)
			print("task scheduled for " + appointment.subject+" at "+ str(appointment.start))

begin = datetime.datetime.now()
end = begin + timedelta(days=1)
print('scheduling zoom meetings for '+begin.strftime('%m/%d/%Y'))
calendar = get_calendar(begin,end)
scheduler = get_scheduler()
schedule_zoom_meetings(calendar,scheduler)
