/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Debug.h
Class Name:  CDebug
Brief:       Debug information functions class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#define  DBLV_IMP		1   // Interface flow 
#define	 DBLV_FUN		2	// Function flow 
#define  DBLV_DATA		3   // Send and Receive data message 
#define	 DBLV_DETAIL  	4	// Internal information   
#define	 DBLV_ERRO 		5	// Error handle information 
#define	 DBLV_USER		6	// Information by definited
#define	 DBLV_RMD		7	// Message of Remote-Debugging tool

#pragma once

class CDebug
{
public:
	CDebug(void);
	~CDebug(void);
	
	static void ShowImpMessage(TCHAR* szMessage);
	static void ShowFunMessage(TCHAR* szMessage);
	static void ShowDataMessage(TCHAR* szMessage);
	static void ShowDetailMessage(TCHAR* szMessage); //show message if Debug or Release
	static void ShowErroMessage(TCHAR* szMessage);
	static void ShowUserMessage(TCHAR* szMessage);
	static void ShowRmdMessage(TCHAR* szMessage);
	
	static void ShowSendMsg(BYTE* pBuf,int nLen);    //Send data( Binary) 
	static void ShowRecMsg(BYTE* pBuf,int nLen);     //receive data( Binary) 

	static void TraceBinData(BYTE* pBuf,int nLen);
	static void LogOut(const TCHAR* szFileName,const BYTE* pBuffer, int nSize, BOOL bHex = TRUE); //output message to a file of txt
private:
	static void ShowBinMessage(BYTE* pBuf,int nLen); //show Binary Message
	static void ShowDebugMessage(int nLevel, TCHAR* szMessage);

public:	
	static bool bShowIMPMsg;	
	static bool bShowERRMsg;
	static bool bShowFUNMsg;	
	static bool bShowDATAMsg;	
	static bool bShowUSERMsg;
	static bool bShowALLMsg;
};
