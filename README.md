# VBProjects

In Progress:  Porting this online.
/Database/ holds all the relevant database code.<br><br>

* Developed using Visual Basic for Applications, Microsoft Access, and MySQL
* Company database for used for transcript production and management
* Saved an average of 15 to 20 hours per week
* Generates 25 to 30 different documents automatically such as price quotes, invoices, cover pages, indexes, and others
* Manages schedule and production workflow
* Auto-imports e-mails from Outlook
* Automatically hyperlinks authority within transcripts via CourtListener API
* Manipulates PDFs to auto-generate bookmarks and create different transcript versions
* Tracks communication and document history of each transcript order
* Manages and plays audio and reporter notes
* Integrated with Office, Acrobat, and WinSCP libraries and several APIs such as CourtListener, Wunderlist, PayPal, OneNote, and others
<br>
This database has been created with concepts from GTD incorporated into its workflow. It uses a strict folder system developed by me to manage all of the related files it deals with.
<br><br>
There is also a working speech recognition component that I have never used because I need something more robust than PocketSphinx; it will do things at the click of a button like record the audio in the proper format, complete the speech recognition, and auto-feed the engine audio & transcripts to make it more accurate. It formats court transcripts into an engine-readable format. A chunk of it was done using batch files, but comes out of VBA & starts from clicking an Access form button.
<br><br>
DB Table Schema @ https://github.com/evoingram/VBProjects/blob/master/database/doc_rptObjects.pdf<br>
DB SQL Queries @ https://github.com/evoingram/VBProjects/blob/master/database/SQL_Queries.pdf<br>
DB Forms Info @ https://github.com/evoingram/VBProjects/blob/master/database/Forms_Info.pdf<br>
DB Relationship Report @ https://github.com/evoingram/VBProjects/blob/master/database/RelationshipReport.pdf<br>
Speech Recognition Batch Files & Components @ https://github.com/evoingram/VBProjects/blob/master/database/speech/<br>
References/Libraries Used @ https://github.com/evoingram/VBProjects/wiki/1.-General----References-Libraries<br>
Batch Files/VBScripts Used @ https://github.com/evoingram/VBProjects/tree/master/database/scripts<br>
