# Logseq Outlook Events

This project uses Python to extract events from an outlook calendar

## Overview

The code hosts a flask server that can be called from a logseq slash command. This server calls python code that connects to an outlook calendar, creates a text output and json output as well as provides the output to the requestor (logseq). That output is input where the slash command was run

Key features:

- Ignore meetings with just the organizer
- Ignore specific meeting subjects
- creates appointment txt and json files that could be fed into an AI/LLM

## Setup

I was not documenting my steps to create this well at all so this is unfinished and would love for someone to test out its creation.

1. Clone this repo
2. rename 'exampleconfig.json' to 'config.json' Configure settings in `config.json` 
3. install dependencies (node.js, pywin32)
4. Schedule `main.py` to run at startup (needs to run as outlook user) (run manually or restart pc)
5. add plugin as an unpacked plugin to logseq

## Usage

- confirm main.py scheduled task is running
- type "/Outlook Events" in logseq.
- a command window should open briefly and input the output from the server.
- only eventsand appointments are pulled in. If you have issues try scheduling a test meeting with other attendees.

## Configuration

Key settings in `config.json`:

- "person_name": First and Last name on the outlook calendar (have not tested other names,calendar names)
- "ignore_canceled": if a meeting is cancelled and this is set to "true" they will not be returned in logseq
- "utilFileLocation": path to the repository example is "c:/logseq-outlook-events"
- "ignore_list": list of subject titles to ignore example: [ "daily team standup","1:1 with manager name" ]

## Future features

- logseq plugin config window
- usage of pip install requirements.txt
- integrate with chat GPT?
- flag for enable creating attendee pages (currently puts attendee names in brackets [[first lane]])

## Bugs

- currently putting all input under a single logseq block

## Contributing

Pull requests welcome!

## License

MIT