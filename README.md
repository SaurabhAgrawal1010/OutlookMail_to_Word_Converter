OutlookMail_to_Word_Converter is very beneficial for a user who needs to extract details from the mail and want to save it in a doc file.

This repository consist of python script which will search for a particular subject of the mail (named as "Mail Subject" in script).
After finding that mail, it will look for a particular keyword in the body of that mail (named as "Example :" in script). As the desired keyword matches, it will extract the words which are present after that keyword and append them in the list.

Then it will put that extracted list in the .doc file, where you want it to be. For this, mergeField functionality of doc file is used. 
Here, the doc file is 'ABC_Template.docx' and mergefield in doc file is <<Example>>

For creating the mergeField:
  Insert --> Quick Parts --> Field --> MergeField
  
This is the generalize repository which needs some modification, according to the user's personal need.

If you want to run this script automatically whenever the mail of a particular subject arrives, copy the code of vbx file in your ThisOutlookSession of OUTLOOK application.
