THIS README ACCOMPANIES A GOOGLE FOLDER WITH A GOOGLE SHEETS FILE IN IT CONTAINING THE FUNCTIONALITY.
IT IS HERE FOR COMPLETENESS ONLY - ACCESS TO THE ACTUAL PROJECT WILL HAVE BEEN INCLUDED IN MY CV.

Bookings & Invoicing System

1. Context & Challenges
2. The Brief
3. Key Operations
4. Access to functionality
4.a Email prerequisites
4.b The resetCols() function

1.
   Context & Challenges:

This was an organic project which grew function by function according to the needs of the business, and so benefitted from a modular approach. The project had to accomplish its functionality within three constraints:


No substantial changes to be made to the spreadsheet or its format by myself
Functionality had to be resilient so some changes could be made by others (particularly, adding or moving columns)
Human-readable formats had to be used in the code (eg, getting sheets by name rather than ID) so that non-technical staff could make basic code edits after my departure

2.
   The Brief:

Given a pre-existing live performance bookings tracker on Google Sheets, I was asked to create an invoicing system that would save the company time - previously invoices were manually filled in, downloaded and attached to emails one by one.

Certain documentation and programming choices were requested: plenty of non-technical comments, use of sheet names instead of IDs, and some general maintenance advice with screenshots for handover (which I can share on request).

3.
   Key Operations:

Know when an invoice is due to be sent
Generate correct invoice reference numbers
Create and export Invoice PDFs
File PDFs into correct folders
Draft emails with invoices attached
Ability retained for user to fill in details & export individual PDFs manually


4.
   Access to Functionality:

Functionality is executed from within the Bookings spreadsheet. Individual invoices can be generated from the invoices sheets, or for more comprehensive functions visit the 'Settings' sheet.

Invoices are exported to subfolders within Invoice System , either to the correct subfolder or to the "Test Invoices" folder if the Debug? box is checked by the user.


Invoice reference numbers will not be generated unless the booking is both confirmed and due for an invoice.


4.a 
Email functionality requires the user to be signed in with a gmail account and for the "from" email in the Settings tab to be a valid alias under the signed in user:

NOTE! If no valid From address is provided, the script will ask for authorisation to use your default user address


4.b
If columns are moved or added on any of the booking information sheets, the resetCols() function will need to be executed. This can be found in the Settings sheet by scrolling down:


BACK TO TOP
