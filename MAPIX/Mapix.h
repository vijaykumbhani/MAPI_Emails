#include <Windows.h>
#include <MAPIDefS.h>
#include <MAPIX.h>

#pragma once
#pragma comment(lib, "mapi32.lib")

#define INBOX L"Inbox"

class Mapix
{
private:
	/* common result */
	HRESULT result;

	/* MAPI Pointer */
	LPMAPISESSION m_lpSession;
	LPMAPITABLE m_lpTable, m_lpInboxTable;
	LPMDB m_lpMsgstore, m_lpInboxMsgStore;
	LPSRowSet m_lpRows;
	LPMAPIFOLDER m_lpFolder;
	SBinary sBin;

	/* some constant */
	CString errorDetails;
	
	bool  selectedFlag;
	unsigned long int  inboxRowCount;
	
	CString senderName,senderEmail, senderSubject, senderBody, SenderReceivedTime;


public:
	/* common count */
	static unsigned long int  cols;

	/* mails details structure */
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

	void  clearCommonObjects();
	void  setError(HRESULT);
	void freeRows(LPSRowSet);
	void  clearAllObjects();
	
	LPMDB getInboxMsgStoreObject();
	LPMAPISESSION getCurrentSession();

	bool login();
	bool logout();
	bool openRootFolder();
	bool openInbox();
	bool openSpecialFolder(CString, SBinary, LPMDB);
	bool getInboxContent(LPMDB);
	bool getInboxMailContent();

	unsigned long int  getRowCountInInboxFolder(LPMDB);

	CString getTimeToFileTimeObjects(FILETIME);
	CString getCurrentError();
	
	/* make structure objects */
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

