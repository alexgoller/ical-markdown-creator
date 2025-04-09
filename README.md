# iCal Markdown Creator

A Python tool that converts iCalendar (iCal) files from shared Outlook calendars into beautifully formatted Markdown files. Perfect for creating readable, shareable calendar views for your team or personal use or sharing it with your favourite LLM for summarization of last weeks calendar.

## Features

- 📅 Extracts events from shared Outlook calendar URLs
- 🔄 Handles both one-time and recurring events
- 🌍 Supports all-day events
- 🕒 Proper timezone handling
- 📝 Generates clean, organized Markdown output
- 📊 Groups events by day for better readability
- 🚫 Removes Zoom and Teams meeting invites

## Requirements

- Python 3.12 or higher
- pip or uv package manager

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/ical-markdown-creator.git
cd ical-markdown-creator
```

2. Create and activate a virtual environment:
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows, use `.venv\Scripts\activate`
```

3. Install dependencies:
```bash
pip install -e .
```

## Usage

Run the script with a shared Outlook calendar URL:

```bash
python ical.py --url "https://outlook.live.com/owa/calendar/your_calendar_url"
```

Optional arguments:
- `--output`: Specify the output file name (default: `weekly_calendar.md`)

Example with custom output file:
```bash
python ical.py --url "https://outlook.live.com/owa/calendar/your_calendar_url" --output "my_calendar.md"
```

## Output Format

The generated Markdown file will contain:
- A header with the current week's date range
- Events grouped by day
- Each event includes:
  - Title
  - Start and end times
  - Location (if available)
  - Organizer (if available)
  - Description (if available)

## Example Output

```markdown
# Calendar Events: April 01 - April 07, 2024

## Monday, April 01
- **Team Meeting** (09:00 - 10:00)
  - Location: Conference Room A
  - Organizer: team@example.com

## Tuesday, April 02
- **Project Review** (14:00 - 15:30)
  - Location: Virtual
  - Description: Quarterly project review meeting
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [icalendar](https://github.com/collective/icalendar)
- Uses [python-dateutil](https://github.com/dateutil/dateutil) for date handling
