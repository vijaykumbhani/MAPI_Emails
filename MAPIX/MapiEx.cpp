// MapiEx.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "Mapix.h"
#include <iostream>
#include <afx.h>

using namespace std;

int _tmain(int argc, _TCHAR* argv[])
{
	// initalize objecects 
	Mapix mapi;

	/* mapi login */
	if(mapi.login())
	{
		cout << "login mapi sucessfully" << endl;
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
				else
					cout << mapi.getCurrentError() << endl;
			}
			else
				cout << mapi.getCurrentError() << endl;
		}
		else
			cout << mapi.getCurrentError() << endl;
	}
	if(mapi.logout())
		cout << "MAPi Logoff Successfully" << endl;
	else
		cout << mapi.getCurrentError() << endl;
	return 0;
}

