#!/usr/bin/env python3
"""
Outlook Calendar Weekly Extractor

This script extracts the current week's calendar events from a shared Outlook calendar URL.
It downloads the iCalendar file from the provided URL, filters events for the current week,
and generates a nicely formatted Markdown file.

Usage:
    python outlook_calendar_extractor.py --url "https://outlook.live.com/owa/calendar/shared_calendar_url" [--output weekly_calendar.md]

Requirements:
    pip install requests icalendar pandas python-dateutil
"""

import argparse
import datetime
import sys
import os
import requests
from icalendar import Calendar
from dateutil.relativedelta import relativedelta
from dateutil.rrule import rrulestr
from dateutil.parser import parse
import pandas as pd

def get_current_week_range():
    """Get the start and end dates for the current week (Monday to Sunday)."""
    today = datetime.datetime.now().date()
    start_of_week = today - datetime.timedelta(days=today.weekday())  # Monday
    end_of_week = start_of_week + datetime.timedelta(days=6)  # Sunday
    
    # Create datetime objects for start and end of day
    start_datetime = datetime.datetime.combine(start_of_week, datetime.time.min)
    end_datetime = datetime.datetime.combine(end_of_week, datetime.time.max)
    
    # Make them timezone-aware (UTC)
    start_datetime = start_datetime.replace(tzinfo=datetime.timezone.utc)
    end_datetime = end_datetime.replace(tzinfo=datetime.timezone.utc)
    
    return start_datetime, end_datetime

def fetch_calendar(url):
    """Fetch the iCalendar file from the given URL."""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Error fetching calendar: {e}", file=sys.stderr)
        sys.exit(1)

def parse_ical_data(ical_data, start_date, end_date):
    """Parse iCalendar data and extract events for the current week."""
    events = []
    
    try:
        calendar = Calendar.from_ical(ical_data)
        
        for component in calendar.walk():
            if component.name == "VEVENT":
                # Extract event details
                summary = str(component.get('summary', 'No Title'))
                description = str(component.get('description', ''))
                location = str(component.get('location', ''))
                organizer = str(component.get('organizer', ''))
                
                # Handle recurring events
                if component.get('rrule'):
                    # Extract recurrence rule
                    rrule_str = component.get('rrule').to_ical().decode('utf-8')
                    dtstart = component.get('dtstart').dt
                    
                    # Convert to UTC if it's a datetime
                    if isinstance(dtstart, datetime.datetime):
                        if dtstart.tzinfo is None:
                            dtstart = dtstart.replace(tzinfo=datetime.timezone.utc)
                    
                    # Create a string representation of the rule
                    rrule_string = f"DTSTART:{dtstart.strftime('%Y%m%dT%H%M%SZ')}\n{rrule_str}"
                    
                    # Make sure start_date and end_date have the same timezone awareness as dtstart
                    start_date_aware = start_date
                    end_date_aware = end_date
                    if isinstance(dtstart, datetime.datetime) and dtstart.tzinfo is not None:
                        if start_date.tzinfo is None:
                            start_date_aware = start_date.replace(tzinfo=datetime.timezone.utc)
                        if end_date.tzinfo is None:
                            end_date_aware = end_date.replace(tzinfo=datetime.timezone.utc)
                    
                    # Get recurring instances in the current week
                    try:
                        rule = rrulestr(rrule_string, forceset=True)
                        instances = rule.between(start_date_aware, end_date_aware, inc=True)
                        
                        for instance in instances:
                            event_start = instance
                            
                            # Calculate event duration
                            duration = None
                            if component.get('dtend'):
                                original_start = component.get('dtstart').dt
                                original_end = component.get('dtend').dt
                                if isinstance(original_start, datetime.datetime) and isinstance(original_end, datetime.datetime):
                                    # Ensure both have same timezone awareness
                                    if original_start.tzinfo is None and original_end.tzinfo is not None:
                                        original_start = original_start.replace(tzinfo=original_end.tzinfo)
                                    elif original_start.tzinfo is not None and original_end.tzinfo is None:
                                        original_end = original_end.replace(tzinfo=original_start.tzinfo)
                                    duration = original_end - original_start
                            
                            # Calculate end time
                            event_end = None
                            if duration:
                                event_end = event_start + duration
                            
                            events.append({
                                'summary': summary,
                                'description': description,
                                'location': location,
                                'organizer': organizer,
                                'start': event_start,
                                'end': event_end,
                                'all_day': False
                            })
                    except Exception as e:
                        print(f"Error processing recurring event: {e}", file=sys.stderr)
                
                # Handle regular events
                else:
                    event_start = component.get('dtstart').dt
                    event_end = component.get('dtend').dt if component.get('dtend') else None
                    
                    # Handle all-day events
                    all_day = False
                    if isinstance(event_start, datetime.date) and not isinstance(event_start, datetime.datetime):
                        all_day = True
                        # Convert to datetime for easier comparison
                        event_start = datetime.datetime.combine(event_start, datetime.time.min)
                        if event_end:
                            event_end = datetime.datetime.combine(event_end, datetime.time.min)
                    
                    # Ensure timezone awareness
                    if isinstance(event_start, datetime.datetime) and event_start.tzinfo is None:
                        event_start = event_start.replace(tzinfo=datetime.timezone.utc)
                    if event_end and isinstance(event_end, datetime.datetime) and event_end.tzinfo is None:
                        event_end = event_end.replace(tzinfo=datetime.timezone.utc)
                    
                    # Ensure all datetimes have timezone info for comparison
                    event_start_compare = event_start
                    event_end_compare = event_end
                    
                    if isinstance(event_start, datetime.datetime):
                        if event_start.tzinfo is None:
                            event_start_compare = event_start.replace(tzinfo=datetime.timezone.utc)
                    
                    if event_end and isinstance(event_end, datetime.datetime):
                        if event_end.tzinfo is None:
                            event_end_compare = event_end.replace(tzinfo=datetime.timezone.utc)
                    
                    # Check if event is in the current week
                    if start_date <= event_start_compare <= end_date or (event_end_compare and start_date <= event_end_compare <= end_date):
                        events.append({
                            'summary': summary,
                            'description': description,
                            'location': location,
                            'organizer': organizer,
                            'start': event_start,
                            'end': event_end,
                            'all_day': all_day
                        })
                    
    except Exception as e:
        print(f"Error parsing calendar: {e}", file=sys.stderr)
        sys.exit(1)
    
    return events

def format_events(events):
    """Format events for display and export."""
    formatted_events = []
    
    for event in events:
        # Format date and time
        start_str = event['start'].strftime('%Y-%m-%d %H:%M') if event['start'] else 'N/A'
        end_str = event['end'].strftime('%Y-%m-%d %H:%M') if event['end'] else 'N/A'
        
        # Format all-day events
        if event.get('all_day', False):
            start_str = event['start'].strftime('%Y-%m-%d') + ' (All day)'
            end_str = event['end'].strftime('%Y-%m-%d') + ' (All day)' if event['end'] else 'N/A'
        
        # Extract email from organizer
        organizer = event['organizer']
        if 'MAILTO:' in organizer:
            organizer = organizer.split('MAILTO:')[1].strip()
        
        formatted_events.append({
            'Summary': event['summary'],
            'Start': start_str,
            'End': end_str,
            'Location': event['location'],
            'Organizer': organizer,
            'Description': event['description']
        })
    
    return formatted_events

def save_to_markdown(events, output_file):
    """Save events to a Markdown file."""
    if not events:
        print("No events found for the current week.")
        return
    
    try:
        # Get current week date range for the title
        today = datetime.datetime.now().date()
        start_of_week = today - datetime.timedelta(days=today.weekday())
        end_of_week = start_of_week + datetime.timedelta(days=6)
        week_range = f"{start_of_week.strftime('%B %d')} - {end_of_week.strftime('%B %d, %Y')}"
        
        with open(output_file, 'w', encoding='utf-8') as md_file:
            # Write header
            md_file.write(f"# Calendar Events: {week_range}\n\n")
            
            # Group events by day
            events_by_day = {}
            for event in events:
                # Extract date from the 'Start' field
                if ' (All day)' in event['Start']:
                    date_str = event['Start'].split(' (All day)')[0]
                else:
                    date_str = event['Start'].split(' ')[0]
                
                if date_str not in events_by_day:
                    events_by_day[date_str] = []
                
                events_by_day[date_str].append(event)
            
            # Sort days
            sorted_days = sorted(events_by_day.keys())
            
            # Write events for each day
            for day in sorted_days:
                # Convert to datetime to get day name
                day_dt = datetime.datetime.strptime(day, '%Y-%m-%d').date()
                day_name = day_dt.strftime('%A')
                
                md_file.write(f"## {day_name}, {day_dt.strftime('%B %d')}\n\n")
                
                # Write events for this day
                for event in events_by_day[day]:
                    # Format time
                    if ' (All day)' in event['Start']:
                        time_str = "All day"
                    else:
                        start_time = event['Start'].split(' ')[1]
                        end_time = event['End'].split(' ')[1] if event['End'] != 'N/A' else 'N/A'
                        time_str = f"{start_time} - {end_time}"
                    
                    # Write event details
                    md_file.write(f"### {event['Summary']}\n\n")
                    md_file.write(f"**Time:** {time_str}\n\n")
                    
                    if event['Location'] and event['Location'] != '':
                        md_file.write(f"**Location:** {event['Location']}\n\n")
                    
                    if event['Organizer'] and event['Organizer'] != '':
                        md_file.write(f"**Organizer:** {event['Organizer']}\n\n")
                    
                    if 'Description' in event and event['Description'] and event['Description'] != '':
                        md_file.write("**Details:**\n\n")
                        # Indent the description
                        # cut details when you read "Join Microsoft Teams Meeting"
                        # or "Join Zoom Meeting"
                        # or "Sie wurden zu einem Zoom-Meeting eingeladen"

                        if "Join Microsoft Teams Meeting" in event['Description']:
                            event['Description'] = event['Description'].split("Join Microsoft Teams Meeting")[0]
                        elif "Join Zoom Meeting" in event['Description']:
                            event['Description'] = event['Description'].split("Join Zoom Meeting")[0]
                        elif "Sie wurden zu einem Zoom-Meeting eingeladen" in event['Description']:
                            event['Description'] = event['Description'].split("Sie wurden zu einem Zoom-Meeting eingeladen")[0]

                        description_lines = event['Description'].split('\n')
                        indented_description = '\n'.join(['    ' + line for line in description_lines])
                        md_file.write(f"{indented_description}\n\n")
                    
                    md_file.write("---\n\n")  # Add separator between events
            
            # Add footer with generation info
            md_file.write(f"\n\n_Generated on {datetime.datetime.now().strftime('%Y-%m-%d at %H:%M')} · {len(events)} events_")
            
        print(f"Events saved to {output_file}")
        
    except Exception as e:
        print(f"Error saving to Markdown: {e}", file=sys.stderr)
        sys.exit(1)

def display_events(events):
    """Display events in a formatted table."""
    if not events:
        print("No events found for the current week.")
        return
    
    # Create DataFrame for pretty printing
    df = pd.DataFrame(events)
    
    # Reorder columns for better display
    columns_order = ['Summary', 'Start', 'End', 'Location', 'Organizer']
    df = df[columns_order]
    
    # Print the table
    print("\nCurrent Week's Events:")
    print("=====================")
    print(df.to_string(index=False))
    print(f"\nTotal events: {len(events)}")

def main():
    """Main function to run the script."""
    parser = argparse.ArgumentParser(description='Extract current week events from an Outlook shared calendar.')
    parser.add_argument('--url', required=True, help='URL of the shared Outlook calendar')
    parser.add_argument('--output', default='weekly_calendar.md', help='Output Markdown file (default: weekly_calendar.md)')
    # stdout is boolean flag
    parser.add_argument('--stdout', action='store_true', help='Output markdown to stdout')
    args = parser.parse_args()
    
    print("Fetching calendar data...", file=sys.stderr)
    calendar_data = fetch_calendar(args.url)
    
    print("Determining current week...", file=sys.stderr)
    start_date, end_date = get_current_week_range()
    print(f"Current week: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}", file=sys.stderr)
    
    print("Extracting events...")
    events = parse_ical_data(calendar_data, start_date, end_date)
    
    print(f"Found {len(events)} events for the current week.", file=sys.stderr)
    formatted_events = format_events(events)
    
    # Display events in the console
    # display_events(formatted_events)
    
    # Save to Markdown
    save_to_markdown(formatted_events, args.output)

    if not os.path.exists(args.output):
        print(f"Output file {args.output} does not exist.", file=sys.stderr)
        sys.exit(1)

    # Print the markdown file to stdout if --stdout flag is set
    if args.stdout:
        with open(args.output, 'r', encoding='utf-8') as md_file:
            content = md_file.read()
            print(content)
            md_file.close()
    else:
        print(f"No output to stdout selected", sys.stderr)
    
if __name__ == "__main__":
    main()
