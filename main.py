import pandas as pd
from icalendar import Calendar

# Function to convert ICS file to DataFrame
def ics_to_dataframe(ics_file):
    events = []
    with open(ics_file, 'r', encoding='utf-8') as file:
        ical = Calendar.from_ical(file.read())
        for component in ical.walk():
            if component.name == "VEVENT":
                # Extract event summary
                summary = str(component.get('summary'))
                
                # Extract email addresses from attendees
                attendees = []
                if component.get('attendee'):
                    for attendee in component.get('attendee'):
                        email = str(attendee).split(':')[1] if ':' in str(attendee) else str(attendee)
                        attendees.append(email)

                event = {
                    'Summary': summary,
                    'Attendees': ', '.join(attendees)  # Join emails into a single string
                }
                events.append(event)
    return pd.DataFrame(events)

# Function to save DataFrame to Excel
def save_to_excel(df, output_file):
    df.to_excel(output_file, index=False, engine='openpyxl')

def main():
    ics_file = 'manuel@dreamven.com.ics'  # Path to your ICS file
    output_file = 'calendar_events_1.xlsx'  # Path for the output Excel file

    df = ics_to_dataframe(ics_file)
    save_to_excel(df, output_file)
    print(f'Successfully converted {ics_file} to {output_file}')

if __name__ == '__main__':
    main()
