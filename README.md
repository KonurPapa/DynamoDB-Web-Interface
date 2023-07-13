# DynamoDB-Web-Interface
Simple code for querying the contents of an AWS DynamoDB table, and displaying them front-end on the web.

About
-
This is in *very* early development, but is meant to be fairly plug-and-play when finished. All you have to do to view is download `web_import.html` and run it from your browser. You can upload it to a webpage to display table contents to your site visitors.

It requires these variables to be manually set in your code:
- Region (where the table is)
- Access Key Id
- Secret Access Key
- DynamoDB table name

These variables are initialized at the top of the `web_import.html`, but in the future they will be saved in a separate init file.

TODOs
- 
Future features to be implemented:

- Store init variables in separate file
- Styling/data formatting
- Scanning/querying for specific items
- Select table from dropdown
