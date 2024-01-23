import win32com.client
import datetime
import json


# Load the configuration file
with open('config.json', 'r') as json_file:
    config = json.load(json_file)
util_file = f'{config["utilFileLocation"]}\\util.py'
json_file = f'{config["utilFileLocation"]}\\appointments.json'
txt_file = f'{config["utilFileLocation"]}\\appointments.txt'
# Create instance of Outlook
Outlook = win32com.client.Dispatch("Outlook.Application")
namespace = Outlook.GetNamespace("MAPI")

# Access the calendar
appointments = namespace.GetDefaultFolder(9).Items

# Restrict to items starting today
appointments.Sort('[Start]')
appointments.IncludeRecurrences = 'True'
today = datetime.date.today()
begin = today.strftime('%m/%d/%Y')
appointments = appointments.Restrict("[Start] >= '" + begin + "' AND [End] <= '" + begin + " 11:59 PM'")

events_list = []
text_file_content = ''

for appointment in appointments:
    # If "ignore_canceled" is true in configuration and "Canceled" is in the subject, skip this appointment
    if config["ignore_canceled"] and 'Canceled' in appointment.Subject:
        continue

    # If appointment subject is in ignore list, skip this appointment
    if appointment.Subject in config["ignore_list"]:
        continue

    req_attendees = [attendee.strip() for attendee in appointment.RequiredAttendees.split(';')] if appointment.RequiredAttendees else []
    opt_attendees = [attendee.strip() for attendee in appointment.OptionalAttendees.split(';')] if appointment.OptionalAttendees else []

    all_attendees = req_attendees + opt_attendees

    # Check the name from the configuration
    if config["person_name"] in all_attendees and len(all_attendees) > 1:
        event_dict = {
            'Subject': appointment.Subject,
            'Start': appointment.Start.strftime('%I:%M %p'),
            'Organizer': appointment.Organizer,
            'Required Attendees': req_attendees,
            'Optional Attendees': opt_attendees,
        }
        events_list.append(event_dict)

        event_string = f'Subject: {appointment.Subject}\n\t- Start: {appointment.Start.strftime("%I:%M %p")}\n'
        event_string += f'\t- Organizer: [[{appointment.Organizer}]]\n\t- Required Attendees: '
        for attendee in req_attendees:
            event_string += f'[[{attendee}]], '
        event_string += f'\n\t- Optional Attendees: '
        for attendee in opt_attendees:
            event_string += f'[[{attendee}]], '

        text_file_content += f'{event_string}\n'
    
# Save the result JSON file
with open(json_file, 'w') as json_file:
    json.dump(events_list, json_file)

# Save the result text file
with open(txt_file, 'w') as text_file:
    text_file.write(text_file_content)

    # Output the text to the console
print(text_file_content)