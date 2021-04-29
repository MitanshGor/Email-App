# Email-App

These Email Desktop software is made for someone who wants to send email  of same content/data and same subject  to multiple organizations, company, people  "AT ONE TIME".

Its really simple to use.
It take max to max 10 clicks to send multiple emails from your DESKTOP.
You can also provide attachemnts to the gmail.
No need to go to gmail tabs to do the same.
Infinite emails can be send at a single time.
Also keeps record of Successful emails Send  and emails left to send .

**Prerequizites required for the Software :**
	
	Python Libraries :
		Numpy should be installed
		Pandas should be installed
		Tinkter should be installed
		scipy should be installed
		smtplib should be installed

	Extra :
		Excel sheet for emails
		Htm file for EmailContent
		Setup Credentials	


Setup :

	There should be an excel sheet with the column name . 
	these column sould contain all the email id to which you want to send the email.

	For making the htm file:
	         Open word docx and jot down the content you want to send.To make it more Viewer-Friendly try to make content designer by giving it colours , suitable text-fonts ,                 text-size you like.After the content is ready go to 'save as' and and save it usingHTM file at your sutable location.

	Credentials should be in the same file as that of python file provided.The first line of this file has "somebody@gmail.com", replace it with the emailid from which you want     to send email.The 2nd line of credentials file should contain the password of the given emailid.


How to use :

	when you open the software 
		1st give path of the EXCEL FILE
		2nd give Path of the HTM FILE
		3rd provide the SUBJECT 
		4th set the COLUMN under which you have your emails in excel
		5th set thee no of ATTACHEMNTS you require
		6th press ATTACH BUTTON ----> if you select attachemts >0 in 5th step than it will shift you to a page where you can add Attachments else email sending process starts
		7th add the attachments one by one 
		8th click on SEND_MAIL button ----> (email sending process starts)
		9th Wait till all mails are send.

