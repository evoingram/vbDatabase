# VBProjects
Learning how to use VB.net instead of VBA; eventually going to port my Access/MySQL db to a web version of some sort.

/Database/ holds all the relevant database code.<br>
/WebApp1/ holds code from me learning ASP.net, learning to port the db to new website


This is the code for my company's database that I use for business ops/transcript management. I wrote it myself (minus jsonconverter/dictionary) & development is still ongoing. It uses an Access database with VBA & MySQL. I save at least 45 mins per transcript order using this database.

This code does so much, I can't really list it all here. A general list is found at https://github.com/evoingram/VBProjects/blob/master/database/AboutDB.bas. I have refactored it since, so the list is a little outdated, and I've also integrated classes to manage all the required data.

Some of the things it does include the following. It automates API interactions with PayPal, Wunderlist, CourtListener for automatically hyperlinking authority in transcripts, & others. It also auto-reads emails from Outlook and 'processes' them. The DB manages my workflow in many ways. It automatically produces a ton of different Office docs according to templates I've created. It manipulates PDFs to add bookmarks, create different transcript versions. It manages & plays audio and reporter notes. It automatically manages my schedule & tells me what i can handle or can't based on tasks it auto-creates & I complete/check off as i go. I can easily send price quotes & download 'new' files from the FTP server when customers upload new audio/files.

This database has been created with concepts from GTD incorporated into its workflow. It uses a strict folder system developed by me to manage all of the related files it deals with.

There is also a working speech recognition component that I have never used because I need something more robust than PocketSphinx; it will do things at the click of a button like record the audio in the proper format, complete the speech recognition, and auto-feed the engine audio & transcripts to make it more accurate. It formats court transcripts into an engine-readable format. A chunk of it was done using batch files, but comes out of VBA & starts from clicking an Access form button.
