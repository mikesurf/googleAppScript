# googleAppScript
A google App Script (JavaScript) to copy and transform entries from one Google sheet to another

## Context
I'm using a Google form written by a third party that sends form fills to my Google Sheet (source sheet). I can't intercept or modify the code that puts the info in that sheet. However, my script, which runs in the Sheet's Google App Script, copies the rows added to the source sheet, then transforms them slightly and puts them into a target sheet. From there I do some other things like have Zapier listen for those newly added rows in the target sheet and take further action.

Also, I have a function _getRichTextValue() which is able to get all rich text value(s) from a cell. This is important because a cell might contain some text and then also "rich text" like a hyperlink (behind the text). And within that function, the call to the getRuns() function is key because it's the thing that accounts for one or more rich text values in the same cell. Otherwise if you were to call getRichTextValue() on a cell it would return the rich text but ONLY if there was one rich text value in that cell.

## How to use
- Place this script as it's own file in the "Files" section of your Google Sheet's App Script. Note that the .gs extension will be added to your filename. I just removed the .js and Google adds .gs when you save the file.

  ![](add-a-file.PNG)

- In my case, I setup a timer to run once every minute since I wanted any new rows added to my source sheet to be moved to my target sheet as quickly as possible. At the moment (9/29/2022) Google Sheets doesn't support a way to "listen" to a sheet for changes so a timer is the best option I could find. 

  ![](timer-example.PNG)