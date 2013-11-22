#include "StdAfx.h"
#include "Mapix.h"


Mapix::Mapix(void)
{
	// mapi handles 
	m_lpSession = NULL;
	m_lpMsgstore = NULL;
	m_lpTable = NULL;
	m_lpRows = NULL;
	m_lpFolder = NULL;
	m_lpInboxMsgStore = NULL;
	m_lpInboxTable = NULL;

	// result flag
	result = NULL;

	// error code 
	errorDetails = L"";

	// get count of mail
	inboxRowCount = 0;

	// to be used static member 
	Mapix::cols = 0;

	/* message elements */
	senderName = L"";
	senderEmail = L"";
	senderBody = L"";
	senderSubject = L"";
	SenderReceivedTime = L"";

	/* selected Flag */
	selectedFlag = FALSE;
	
}


Mapix::~Mapix(void)
{

}

/* mapi login */
bool Mapix::login()
{
	// Mapix initialize
	MAPIINIT_0 MAPIInit = {MAPI_INIT_VERSION, MAPI_MULTITHREAD_NOTIFICATIONS};
	result = MAPIInitialize(&MAPIInit);
	if(result == S_OK)
	{
		result = MAPILogonEx(NULL, L"", L"", MAPI_ALLOW_OTHERS|MAPI_USE_DEFAULT|MAPI_EXTENDED|MAPI_NEW_SESSION, &m_lpSession);
		if(result != S_OK)
		{
			setError(result);
			return 0;
		}
		else
			return 1;
	}
	else
	{
		setError(result);
		return 0;
	}
}

/* static cols count*/
unsigned long int Mapix::cols = 0;

/* opern root folder */
bool Mapix::openRootFolder()
{
	if(m_lpSession)
	{
		// open msg store table that open mailbox in outlook
		clearCommonObjects();
		result = m_lpSession->GetMsgStoresTable(0, &m_lpTable);
		if(result != S_OK)
		{
			m_lpTable->Release();
			setError(result);
			return 0;
		}
		else
		{
			const int nProperties = 3;
			SizedSPropTagArray(nProperties, Column) = {nProperties, {PR_ENTRYID, PR_DEFAULT_STORE, PR_DISPLAY_NAME}};
			result = m_lpTable->SetColumns((LPSPropTagArray)&Column, 0);
			if(result == S_OK)
			{
				while(m_lpTable->QueryRows(1,0, &m_lpRows) == S_OK)
				{
					if(m_lpRows->cRows != 1)
						break;
					else
					{
						if(m_lpRows->aRow[0].lpProps[1].Value.b)
							break;
					}
				}
			
				if(m_lpRows->aRow[0].lpProps[1].Value.b)
				{
					result = m_lpSession->OpenMsgStore(NULL, m_lpRows->aRow[0].lpProps[0].Value.bin.cb,(LPENTRYID)m_lpRows->aRow[0].lpProps[0].Value.bin.lpb, NULL, MAPI_BEST_ACCESS|MDB_NO_DIALOG, &m_lpMsgstore);
					if(result != S_OK)
					{
						setError(result);
						return 0;
					}
					else
					{
						// release table objects
						clearCommonObjects();

						ULONG cbEntryID = 0; 
						LPENTRYID lpEntryID = NULL; 
						ULONG ulObjectType; 
						result = m_lpMsgstore->OpenEntry(cbEntryID, lpEntryID, NULL, MAPI_MODIFY|MAPI_BEST_ACCESS, &ulObjectType, (LPUNKNOWN*)&m_lpFolder);
						if(result != S_OK)
						{
							setError(result);
							return 0;
						}
						else
							return 1;
					}
				}
				else
					return 0;
			}
			else
			{
				setError(result);
				return 0;
			}
		}
	}
	else
		return 0;
}

/* open inbox */
bool Mapix::openInbox()
{
	if(m_lpSession && m_lpMsgstore && m_lpFolder)
	{
		clearCommonObjects();
		
		result = m_lpFolder->GetHierarchyTable(NULL, &m_lpTable);
		if(result != S_OK)
		{
			setError(result);
			return 0;
		}
		else
		{
			const int nProperties = 2;
			SizedSPropTagArray(nProperties, Column) = {nProperties, {PR_DISPLAY_NAME, PR_ENTRYID}};
			result = m_lpTable->SetColumns((LPSPropTagArray)&Column, 0);
			if(result == S_OK)
			{
				while(m_lpTable->QueryRows(1,0, &m_lpRows) == S_OK)
				{
					if(m_lpRows->cRows != 1)
						break;
					else
					{
						selectedFlag = openSpecialFolder(INBOX, m_lpRows->aRow[0].lpProps[1].Value.bin, m_lpMsgstore);
						if(selectedFlag)
						{
							selectedFlag = FALSE;
							clearCommonObjects();
							return 1;
						}
					}
				}
			}
			else
			{
				setError(result);
				return 0;
			}
		}
	}
	else
		return 0;
	return 0;
}

/* open special folder */
bool Mapix::openSpecialFolder(CString folderName, SBinary bin, LPMDB msgStore)
{
	
	LPMAPITABLE table = NULL;
	LPMAPIFOLDER m_folder= NULL;
	ULONG objectType = NULL;
	LPSRowSet pRows = NULL;

	result = msgStore->OpenEntry(bin.cb, (LPENTRYID)bin.lpb, NULL,  MAPI_MODIFY|MAPI_BEST_ACCESS, &objectType, (LPUNKNOWN*)&m_folder);

	if(result != S_OK)
	{
		setError(result);
		return 0;
	}
	else
	{	
		result = m_folder->GetHierarchyTable(NULL, &table);
		if(result == S_OK)
		{	
			const int nProperties = 2;
			SizedSPropTagArray(nProperties, Column) = {nProperties, {PR_DISPLAY_NAME, PR_ENTRYID}};
			result = table->SetColumns((LPSPropTagArray)&Column, 0);
			if(result == S_OK)
			{
				while(table->QueryRows(1,0, &pRows) == S_OK)
				{
					if(pRows->cRows != 1)
						break;
					else
					{
						CString nameOfFolder( pRows->aRow[0].lpProps[0].Value.lpszW);
						if(nameOfFolder == folderName)
						{
							sBin = pRows->aRow[0].lpProps[1].Value.bin;
							m_lpInboxMsgStore = msgStore;
							m_folder->Release();
							table->Release();
							pRows = NULL;
							return 1;
						}

						// open enumarate folder
						openSpecialFolder(folderName, pRows->aRow[0].lpProps[1].Value.bin, msgStore);
					}
				}
			}
			else
			{
				setError(result);
				return 0;
			}
		}
		else
		{
			setError(result);
			return 0;
		}
	}
	return 0;
}

/* get row count in inbox */
unsigned long int Mapix::getRowCountInInboxFolder(LPMDB m_InboxMsgStoreFolder = NULL)
{
	if(m_InboxMsgStoreFolder == NULL)
	{
		if(m_lpInboxMsgStore)
		{
			m_InboxMsgStoreFolder = m_lpInboxMsgStore;
			clearCommonObjects();
		}
	}
	if(m_InboxMsgStoreFolder)
	{
		ULONG dwObjectType;
		LPMAPIFOLDER m_Folder = NULL;
		result = m_InboxMsgStoreFolder->OpenEntry(sBin.cb, (LPENTRYID)sBin.lpb, NULL, MAPI_MODIFY|MAPI_BEST_ACCESS, &dwObjectType, (LPUNKNOWN*)&m_Folder);
		if(result == S_OK)
		{
			result = m_Folder->GetContentsTable(MAPI_UNICODE|MAPI_DEFERRED_ERRORS, &m_lpTable);
			if(result != S_OK)
			{
				setError(result);
				return 0;
			}
			else
			{
				m_lpTable->GetRowCount(0,&inboxRowCount);
				clearCommonObjects();
				m_Folder->Release();
				return inboxRowCount;
			}
		}
		else
		{
			setError(result);
			return 0;
		}
	}
	else
		return 0;
}

/* get inbox all emails in inbox folder */
bool Mapix::getInboxContent(LPMDB m_InboxMsgStore = NULL)
{
	if(m_InboxMsgStore == NULL)
	{
		if(m_lpInboxMsgStore)
		{
			m_InboxMsgStore = m_lpInboxMsgStore;
			clearCommonObjects();
		}
		else
			return 0;
	}
	if(m_InboxMsgStore)
	{
		ULONG dwObjectType;
		LPMAPIFOLDER m_Folder = NULL;
		result = m_InboxMsgStore->OpenEntry(sBin.cb, (LPENTRYID)sBin.lpb, NULL, MAPI_MODIFY|MAPI_BEST_ACCESS, &dwObjectType, (LPUNKNOWN*)&m_Folder);
		if(result == S_OK)
		{
			ULONG rowCount = getRowCountInInboxFolder(m_InboxMsgStore);
			result = m_Folder->GetContentsTable(MAPI_UNICODE|MAPI_DEFERRED_ERRORS, &m_lpTable);
			if(result != S_OK)
			{
				setError(result);
				return 0;
			}
			else
			{	
				contentOfMessage = new MailContent[rowCount+1];
				const int nProperties = 6;
				int i = 0;
				SizedSPropTagArray(nProperties, Column) = {nProperties, PR_ENTRYID, PR_SENDER_NAME, PR_SENDER_EMAIL_ADDRESS, PR_BODY, PR_SUBJECT, PR_MESSAGE_DELIVERY_TIME};
				result = m_lpTable->SetColumns((LPSPropTagArray)&Column, 0);
				if(result == S_OK)
				{
					while(m_lpTable->QueryRows(1, 0, &m_lpRows) == S_OK)
					{
						if(m_lpRows->cRows != 1)
							break;
						else
						{
							if(m_lpRows->aRow[0].lpProps[1].Value.lpszW)
							{
								contentOfMessage[i].senderName = m_lpRows->aRow[0].lpProps[1].Value.lpszW;
								contentOfMessage[i].SenderReceivedTime = getTimeToFileTimeObjects(m_lpRows->aRow[0].lpProps[5].Value.ft);
							}
							if(m_lpRows->aRow[0].lpProps[2].Value.lpszW)
								contentOfMessage[i].senderEmail = m_lpRows->aRow[0].lpProps[2].Value.lpszW;
							if(m_lpRows->aRow[0].lpProps[3].Value.lpszW)
								contentOfMessage[i].senderBody = m_lpRows->aRow[0].lpProps[3].Value.lpszW;
							if(m_lpRows->aRow[0].lpProps[4].Value.lpszW)
								contentOfMessage[i].senderSubject = m_lpRows->aRow[0].lpProps[4].Value.lpszW;
						}
						i++;
					}
				}
				else
				{
					setError(result);
					return 0;
				}
			}
		}
		else
		{
			setError(result);
			return 0;
		}
	}
	else
		return 0;
	return 1;
}

/* convert fite time objects string time format */
CString Mapix::getTimeToFileTimeObjects(FILETIME ft)
{
	CString timeAndDateInString;
	TCHAR szTime[256];
	FILETIME localFileTime;
	SYSTEMTIME tm;
	FileTimeToLocalFileTime(&ft, &localFileTime);
	FileTimeToSystemTime(&localFileTime, &tm);
	LPCTSTR szFormat=_T("MM/dd/yyyy hh:mm:ss tt");
	GetDateFormat(LOCALE_SYSTEM_DEFAULT, 0, &tm, szFormat, szTime, 256);
	GetTimeFormat(LOCALE_SYSTEM_DEFAULT, 0, &tm, szTime, szTime, 256);
	timeAndDateInString = szTime;
	return timeAndDateInString;
}

/* get mail content from inbox */
bool Mapix::getInboxMailContent()
{
	setSenderName(contentOfMessage[Mapix::cols].senderName);
	setSenderEmail(contentOfMessage[Mapix::cols].senderEmail);
	setSenderBody(contentOfMessage[Mapix::cols].senderBody);
	setSenderSubject(contentOfMessage[Mapix::cols].senderSubject);
	setSenderTime(contentOfMessage[Mapix::cols].SenderReceivedTime);

	if(Mapix::cols == inboxRowCount)
	{
		delete [] contentOfMessage;
		Mapix::cols = 0;
		return 0;
	}
	else
		Mapix::cols++;
		
	return 1;
}

/* get inbox msg store objects */
LPMDB Mapix::getInboxMsgStoreObject()
{
	return m_lpInboxMsgStore;
}

/* get current session */
LPMAPISESSION Mapix::getCurrentSession()
{
	return m_lpSession;
}

/* clear table and rowset */
void Mapix::clearCommonObjects()
{
	if(m_lpTable || m_lpRows)
	{
		m_lpTable->Release();
		m_lpTable = NULL;
		if(m_lpRows)
		{
			freeRows(m_lpRows);
			m_lpRows = NULL;
		}
	}
}

/* free rows */
void Mapix::freeRows(LPSRowSet nRows)
{
	if(nRows) 
	{
		for(ULONG i=0;i<nRows->cRows;i++) 
		{
			MAPIFreeBuffer(nRows->aRow[i].lpProps);
		}
		MAPIFreeBuffer(nRows);
	}
}

/* log out */
bool Mapix::logout()
{
	if(m_lpSession)
	{
		clearCommonObjects();
		clearAllObjects();
		result = m_lpSession->Logoff(NULL, MAPI_LOGOFF_SHARED, 0);
		if(result == S_OK)
			m_lpSession->Release();
		MAPIUninitialize();
		return 1;
	}
	return 0;
}

/* set sender name */
void Mapix::setSenderName(CString name)
{
	senderName = name;
}

/* set sender email */
void Mapix::setSenderEmail(CString email)
{
	senderEmail = email;
}

/* set sender body */
void Mapix::setSenderBody(CString body)
{
	senderBody = body;
}

/* set sender body */
void Mapix::setSenderSubject(CString subject)
{
	senderSubject = subject;
}

/* set sender time */
void Mapix::setSenderTime(CString receivedTime)
{
	SenderReceivedTime = receivedTime;
}


/* get current error */
CString Mapix::getCurrentError()
{
	return errorDetails;
}

/* get sender name */
CString Mapix::getSenderName()
{
	return senderName;
}

/*get sender email */
CString Mapix::getSenderEmail()
{
	return senderEmail;
}

/* get sender body */
CString Mapix::getSenderBody()
{
	return senderBody;
}

/* get sender subject */
CString Mapix::getSenderSubject()
{
	return senderSubject;
}

/* get sender time */
CString Mapix::getSenderTime()
{
	return SenderReceivedTime;
}


/* clear all objects */
void Mapix::clearAllObjects()
{
	if(m_lpFolder)
		m_lpFolder->Release();
	if(m_lpMsgstore)
		m_lpMsgstore->Release();
	if(m_lpInboxTable)
		m_lpInboxTable->Release();
}

/* set curerent error */
void Mapix::setError(HRESULT errorCode)
{
	switch(errorCode)
	{
		case MAPI_E_LOGON_FAILED:
			errorDetails = L"The logon was unsuccessful, either because one or more of the parameters to MAPILogonEx were invalid or because there were too many sessions open already.";
			break;
		case MAPI_E_TIMEOUT:
			errorDetails = L"MAPI serializes all logons through a mutex. This is returned if the MAPI_TIMEOUT_SHORT flag was set and another thread held the mutex.";
			break;
		case MAPI_E_USER_CANCEL:
			errorDetails = L"The user canceled the operation, typically by clicking the Cancel button in a dialog box.";
			break;
		case MAPI_E_BAD_CHARWIDTH:
			errorDetails = L"The MAPI_UNICODE flag was set and the session does not support Unicode.";
			break;
		case MAPI_E_NO_ACCESS:
			errorDetails = L"An attempt was made to access a message store for which the user has insufficient permissions.";
			break;
		case MAPI_E_NOT_FOUND:
			errorDetails = L"The message store indicated by lpEntryID does not exist.";
			break;
		case MAPI_E_UNKNOWN_CPID:
			errorDetails = L"The server is not configured to support the client's code page.";
			break;
		case MAPI_E_UNKNOWN_LCID:
			errorDetails = L"The server is not configured to support the client's locale information.";
			break;
		case MAPI_W_ERRORS_RETURNED:
			errorDetails = L"The call succeeded, but the message store provider has error information available. When this warning is returned, the call should be handled as successful.";
			break;
		case MAPI_E_UNKNOWN_ENTRYID:
			errorDetails = L"The entry identifier passed in the lpEntryID parameter is in an unrecognizable format";
			break;
	}
}