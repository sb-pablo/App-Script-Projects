# App-Script-Projects

This Script parses through a list of companies and compares them to an existing account list. It then populates a new tab in the Google Sheet with any matches it finds. 

A couple of things to note. 

1. Make sure that the columns match. From left to right it should be "rep_accounts, LAST NAME, FIRST NAME, company_names, GITHUB HANDLE, JOB TITLE, COUNTRY, TICKET TYPE, ATTENDEE TYPE, REGISTERED DATE""
2. For rep_accounts ensure that you put the main company name there. For example, if an account is called Vanguard_Parent and there are multiple different accounts similar to it like Vanguard_Server, just type Vanguard.
3. Feel free to use GitHub Copilot for troubleshooting if you have any questions on the code. Its free for all GitHub Employees and pretty cool! 


How to use this. 

1. Make a copy of my current sheet and use that copy for steps 3-10. (https://docs.google.com/spreadsheets/d/1pb1FooEDZ9zqdZ6FhfgEf9-rXUNpXJPGz3n1FCJjX_Y/edit?usp=sharing)
2. Click on extensions
3. Click on App Script
4. Accept the T&C's if you haven't already
5. Paste the "Code.gs" file into the files section
6. Click the Save icon.
7. Refresh your google sheet.
8. Click on the Custom Menu tab
9. Select Match and Copy Rows.
