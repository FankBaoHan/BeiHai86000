/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Debug.cpp
Class Name:  CDebug
Brief:       Debug information functions class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#include "StdAfx.h"
#include "debug.h"

const int WM_MESSAGE_OUTPUT = WM_USER + 50;

bool CDebug::bShowIMPMsg;
bool CDebug::bShowERRMsg;
bool CDebug::bShowFUNMsg;
bool CDebug::bShowDATAMsg;
bool CDebug::bShowUSERMsg;
bool CDebug::bShowALLMsg;

CDebug::CDebug(void)
{
	bShowIMPMsg = 0;	
	bShowERRMsg = 0;
	bShowFUNMsg = 0;	
	bShowDATAMsg = 0;	
	bShowUSERMsg = 0;
	bShowALLMsg = 0;
}

CDebug::~CDebug(void)
{
}

/****************************************************************************
*   Name
ShowDebugMessage
*	Type
private
*	Function
the final function of Showing  Message with different level, it can show 127 TCHASs at most
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to display message
[in]   the level of message
*****************************************************************************/
void CDebug::ShowDebugMessage(int nLevel, TCHAR* szMessage)
{
	HWND hMessWnd = ::FindWindow( _T("kingmess"), NULL);
    TCHAR szMess[128];//
	memset(szMess, 0, sizeof(szMess));
    
	STR  messAtom;
	LARGE_INTEGER litmp ; 
    LONGLONG QPart = 0; 
    double dfFreq = 0;
	double dfTim = 0; 
    QueryPerformanceFrequency(&litmp); 
	dfFreq = (double)litmp.QuadPart; 
	QueryPerformanceCounter(&litmp);	 
	QPart = litmp.QuadPart;	
	dfTim = QPart / dfFreq*1000; 
	CString  strPerformance = _T("");
	strPerformance.Format(_T("%.05f "), dfTim);

	if( hMessWnd)
	{	
		_tcsncpy( szMess, strPerformance,min(127,  _tcslen(strPerformance)) );
		_tcsncpy( szMess+_tcslen(strPerformance),  szMessage,min(127-_tcslen(strPerformance), _tcslen(szMessage)) );  //           
		if(_tcslen(szMessage + strPerformance) >= (127))//
			szMess[127]=0;//
		else
			szMess[_tcslen(szMessage + strPerformance)] = 0;
		messAtom =CHelper::KvGlobalAddAtom( szMess );
        
		ASSERT( messAtom );
		unsigned long messId = WM_MESSAGE_OUTPUT;
		::SendMessage( hMessWnd, WM_MESSAGE_OUTPUT, (WPARAM)messAtom, 0 );
        CHelper::KvGlobalDeleteAtom( messAtom );
		
	}
}


/****************************************************************************
*   Name
		ShowImpMessage
*	Type
		public
*	Function
		Show Interface flow Message
*	Return Value
		null
*	Parameters
		szMessage 
			[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowImpMessage(TCHAR* szMessage)
{
	if((bShowIMPMsg) || (bShowALLMsg))
	{
		ShowDebugMessage(DBLV_IMP,szMessage);
	}	
}

/****************************************************************************
*   Name
		ShowFunMessage
*	Type
		public
*	Function
		Show Function flow Message
*	Return Value
		null
*	Parameters
		szMessage 
			[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowFunMessage(TCHAR* szMessage)
{
	if((bShowFUNMsg) || (bShowALLMsg))
	{
		ShowDebugMessage(DBLV_FUN,szMessage);
	}
}


/****************************************************************************
*   Name
		ShowDataMessage
*	Type
		public
*	Function
		Show Send and Receive data  Message in ascii
*	Return Value
		null
*	Parameters
		szMessage 
			[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowDataMessage(TCHAR* szMessage)
{
	if((bShowDATAMsg) || (bShowALLMsg))
	{
		ShowDebugMessage(DBLV_DATA,szMessage);
	}
}

/****************************************************************************
*   Name
ShowDetailMessage
*	Type
public
*	Function
Show detail message,these message will be shown in the kingmess window
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowDetailMessage(TCHAR* szMessage)
{
	ShowDebugMessage(DBLV_DETAIL,szMessage);
}


/****************************************************************************
*   Name
		ShowErroMessage
*	Type
		public
*	Function
		Show Error handle message
*	Return Value
		null
*	Parameters
		szMessage 
			[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowErroMessage(TCHAR* szMessage)
{
	if((bShowERRMsg) || (bShowALLMsg))
	{
		ShowDebugMessage(DBLV_ERRO,szMessage);
	}
}

//
/****************************************************************************
*   Name
		ShowUserMessage
*	Type
		public
*	Function
		Show Information by self-definited
*	Return Value
		null
*	Parameters
		szMessage 
			[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowUserMessage(TCHAR* szMessage)
{
	if((bShowUSERMsg) || (bShowALLMsg))
	{
		ShowDebugMessage(DBLV_USER,szMessage);
	}
}

/****************************************************************************
*   Name
ShowRmdMessage
*	Type
public
*	Function
Show remote debuging message
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to display message
*****************************************************************************/
void CDebug::ShowRmdMessage(TCHAR* szMessage)
{
	ShowDebugMessage(DBLV_RMD,szMessage) ;
}

/****************************************************************************
*   Name
ShowRmdMessage
*	Type
private
*	Function
Show message of Hex
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to data buffer
[in]   the length of data buffer 
*****************************************************************************/
void CDebug::ShowBinMessage(BYTE* pBuf,int nLen)
{
	CString szVal ;
	TCHAR   szBuf[500]={0} ;
	
	for (int i=0 ; i<nLen ; i++)
	{
		_stprintf(szBuf,  _T("%02X"), pBuf[i]) ;
		szVal = szVal + szBuf + ' ' ;
	}
	
	ShowDetailMessage(szVal.GetBuffer(0)) ;
}

/****************************************************************************
*   Name
ShowSendMsg
*	Type
private
*	Function
Show sending message 
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to data buffer
[in]   the length of data buffer 
*****************************************************************************/
void CDebug::ShowSendMsg(BYTE* pBuf,int nLen)
{
	if((bShowDATAMsg) || (bShowALLMsg))
	{
		CString szVal0 = _T(" SendData:") ;
		CString szVal = _T("") ;

		TCHAR   szBuf[500]={0} ;  //
		for (int i=0 ; i<nLen ;i++)
		{
			_stprintf(szBuf,  _T("%02X"), pBuf[i]) ;
			szVal = szVal + szBuf + ' ' ;
		}
		ShowDetailMessage(szVal0.GetBuffer(0)) ;
		ShowDetailMessage(szVal.GetBuffer(0)) ;
	}
}

/****************************************************************************
*   Name
ShowRecMsg
*	Type
private
*	Function
Show receiving message 
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to data buffer
[in]   the length of data buffer 
*****************************************************************************/
void CDebug::ShowRecMsg(BYTE* pBuf,int nLen)
{
	if((bShowDATAMsg) || (bShowALLMsg))
	{
	CString szVal0 = _T(" RecData:") ;
	CString szVal = _T("") ;

	TCHAR   szBuf[500]={0} ;  //
	for (int i=0; i<nLen; i++)
	{
		_stprintf(szBuf,  _T("%02X"), pBuf[i]) ;
		szVal = szVal + szBuf + ' ' ;
	}
	ShowDetailMessage(szVal0.GetBuffer(0)) ;
	ShowDetailMessage(szVal.GetBuffer(0)) ;
	}
}

/****************************************************************************
*   Name
TraceBinData
*	Type
public
*	Function
Show message to ouput window if vc.net
*	Return Value
null
*	Parameters
szMessage 
[in]   the pointer to data buffer
[in]   the length of data buffer 
*****************************************************************************/
void CDebug::TraceBinData(BYTE* pBuf,int nLen)
{
	if(pBuf == NULL)
	{
		return;
	}
	for ( int i= 0; i<nLen; i++)
	{
		TRACE1("%02X ", pBuf[i]) ;
	}
	TRACE(_T("\n")) ;
}

/****************************************************************************
*   Name
LogOut
*	Type
public
*	Function
output message to a file of txt, the path of the file is the KV Path
*	Return Value
null
*	Parameters
szFileName [in] the name of message file
pBuffer    [in] the pointer to data buffer
nSize      [in] the length of data you want to output 
bHex       [in] the flag of data    TRUE:Hex  FALSE:Ascii
*****************************************************************************/
void CDebug::LogOut(const TCHAR* szFileName, const BYTE* pBuffer, int nSize, BOOL bHex)
{
	// check the function parameters
	if(szFileName == NULL || pBuffer == NULL)
	{
		return;
	}
	// get the path of kingview
	CString strFilePath;
	TCHAR szKVPathName[256];
	memset(szKVPathName, 0, sizeof(szKVPathName));
	HMODULE hModule = GetModuleHandle(NULL);
	DWORD dwRes = GetModuleFileName(hModule, szKVPathName, sizeof(szKVPathName));

	if(dwRes  == 0)
	{
		return;
	}
	else
	{
		CString strKVPathName(szKVPathName);

		int nIndex = strKVPathName.ReverseFind( _T('\\'));

		strFilePath = strKVPathName.Left(nIndex + 1);

		strFilePath += szFileName;
	}

	// open file 
	FILE * pRecFile = _tfopen(strFilePath, _T("a"));
	if(pRecFile == NULL)
	{
		return;
	}
	else
	{
		SYSTEMTIME  SysTime;
		//get the local time 
		GetLocalTime(&SysTime);
		TCHAR szTimeBuf[256];
		memset(szTimeBuf, 0, sizeof(szTimeBuf));
		int nNum = _stprintf(szTimeBuf, _T("\r\n%02d.%02d.%02d %02d.%02d.%02d.%03d:\n"),
			SysTime.wYear,
			SysTime.wMonth,
			SysTime.wDay,
			SysTime.wHour,
			SysTime.wMinute,
			SysTime.wSecond,
			SysTime.wMilliseconds);
		//output the time to file
		_ftprintf(pRecFile, _T("%s"), szTimeBuf);
		
		if(bHex == TRUE)
		{
			//output pBuffer to file in hex
			for(int nIndex = 0; nIndex < nSize; ++nIndex)
			{
				_ftprintf(pRecFile, _T("%02x  "), pBuffer[nIndex]);
			}
		}
		else
		{
			//output pBuffer to file in ascii
			for(int nIndex = 0; nIndex < nSize; ++nIndex)
			{
				_ftprintf(pRecFile, _T("%c"), pBuffer[nIndex]);
			}
		}
	}
	// close the file
	fclose(pRecFile);
}
