[![Maintainability](https://api.codeclimate.com/v1/badges/06260d9e8729d5d17f2a/maintainability)](https://codeclimate.com/github/evoingram/vbDatabase/maintainability)

# Visual Basic Database Software

## Project Description

Developed pre-Lambda School.  Deployed version is offline.  Porting this to a [React/Redux version](https://github.com/evoingram/webapp-frontend) with a [back end](https://github.com/evoingram/webapp-backend) via Node, Express, and PostgreSQL.  Back end is largely completed; front end in planning stages.

This is my company's VB database, which is used for transcript production and workflow management.  The VB database does everything from scheduling to automated shorthand to automated hyperlinking and document formatting, shipping, production, management of company financials, and other business operations.  Copyright 2020 Erica Ingram.

## Key Features

- Live solo project
- Integrates GTD principles into workflow
- Generates 25 to 30 different documents automatically such as price quotes, invoices, cover pages, indexes, and others
- Manages schedule and production workflow
- Auto-imports electronic communication
- Automatically hyperlinks authority within transcripts via CourtListener API
- Manipulates PDFs to auto-generate bookmarks and create different transcript versions
- Auto-creates and formats Word versions
- Tracks communication and document history of each transcript order
- Manages and plays audio and reporter notes
- Integrated with Office, Acrobat, and WinSCP libraries and several APIs such as CourtListener, PayPal, OneNote, and others
- speech recognition

## Tech Stack

Software runs in `Microsoft Access` and built using:

- [Visual Basic](https://github.com/dotnet/vblang): Visual Basic is an approachable language with a simple syntax for building type-safe, object-oriented apps.
- [SQL](https://en.wikipedia.org/wiki/SQL)a domain-specific language used in programming and designed for managing data held in an RDBMS or stream processing in an RDSMS

## APIs

- [CourtListener](http://courtlistener.com/):  assists in automation of authority hyperlinking.
- [Office](https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/):  document creation, formatting, and management.
- [PayPal](https://developer.paypal.com/home/):  payment processing and management.
- [Wunderlist](https://developer.wunderlist.com/):  to-do list management.
   
## Testing

- Only manual testing has been conducted.

## Documentation

Check out the [Wiki](https://github.com/evoingram/vbDatabase/wiki) for more info on general production workflow, file and ffolder organization, database-related information, templates, classes, modules and various functions within the software.
