#include <Windows.h>
#include <MAPIDefS.h>
#include <MAPIX.h>
#include <afx.h>

#pragma once

#pragma comment(lib, "mapi32.lib")

class Mapix
{
private:
	
	LPMAPISESSION m_lpSession;
	
	LPMDB m_lpMsgstore;
	LPMDB m_lpInboxMsgStore;

	LPMAPITABLE m_lpTable;
	LPMAPITABLE m_inboxTable;

	LPMAPIFOLDER m_lpFolder;

	LPSRowSet m_lpRows;

	HRESULT result;

	SBinary sBin;

	CString errorDetails;
	
	bool selectFlag;
	ULONG inboxRowCount;
	CString senderName,senderEmail, senderSubject, senderBody, SenderReceivedTime;

public:
	static int cols;

	typedef struct 
	{
		CString senderName;
		CString senderEmail;
		CString senderSubject;
		CString senderBody;
		CString SenderReceivedTime;
	} MailContent;

	Mapix(void);
	~Mapix(void);

	void clearCommonObjects();
	void freeRows(LPSRowSet);
	void setError(HRESULT);

	bool login();
	bool logout();
	bool openRootFolder();
	bool openInbox();
	bool openSpecialFolder(CString, SBinary, LPMDB);
	bool getInboxContent(LPMDB);
	bool getInboxMailContent();

	CString getCurrentError();
	LPMDB getInboxMsgStoreObject();
	ULONG getRowCountInInboxFolder(LPMDB);
	CString getTimeToFileTimeObjects(FILETIME);

	
	MailContent *contentOfMessage;

	/* get personal mail content */
	CString getSenderName();
	CString getSenderEmail();
	CString getSenderBody();
	CString getSenderSubject();
	CString getSenderTime();

	/* set personal mail content */
	void setSenderName(CString);
	void setSenderEmail(CString);
	void setSenderBody(CString);
	void setSenderSubject(CString);
	void setSenderTime(CString);

};

