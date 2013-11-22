MAPI_Emails
===========

To get emails from Outlook 


* This can be used only C++/MFC for Windows.
	
* To get Emails from Outlook fast and quick

features:

	1. easy to use
	2. fast and quick 
	3. no limit to get emails from Outlook
	
	
Functions:

Note: Here, all fuctions can be working in windows

	1. bool login():
		login to mapi api
		
	2. bool openRootFolder()
		to open root folder
		
	3. bool openInbox()
		to open inbox folder 
		
	4. bool getInboxContent()
		to get all inbox mails
	
	5. bool getInboxMailContent()
		to get all emails one by one 
		ex : senderName, senderEmail, senderSubjects, senderBody, senderReceivedTime

	6. CString getSenderName()
		get sender name

	7. CString getSenderEmail()
		get sender email

	8. CString getSenderSubject()
		get sender subjects 
	
	9. CString getSenderBody()
		get sender body

	10. CString getSenderTime()
		get sender received time 
		
		
Examples :


	 // initalize objects 
        Mapix mapi;

        /* mapi login */
        if(mapi.login())
        {
                cout << "login mapi successfully" << endl;
                /* opern root folder */
                if(mapi.openRootFolder())
                {
                        cout << "open root folder" <<endl;
                        /* open inbox */
                        if(mapi.openInbox())
                        {
                                cout << "open inbox" << endl;
                                /* open inbox email content */
                                if(mapi.getInboxContent(NULL))
                                {
                                        cout <<"get inbox mail" << endl;
                                        /* get email one by one */
                                        while(mapi.getInboxMailContent())
                                        {
                                                cout<<mapi.getSenderName()<<endl;
                                                cout<<mapi.getSenderEmail()<<endl;
                                                cout<<mapi.getSenderSubject()<<endl;
                                                cout<<mapi.getSenderBody()<<endl;
                                                cout<<mapi.getSenderTime()<<endl;
                                        }
                                }
                        }
                }
        }
        if(mapi.logout())
        	cout << "mapi logout successfully";

	
