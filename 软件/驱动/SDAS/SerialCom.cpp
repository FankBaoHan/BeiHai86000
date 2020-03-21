/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   SerialCom.cpp
Class Name:  CSerialCom
Brief:       Serial Communication class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#include "StdAfx.h"
#include "serialcom.h"

#ifdef wReserved
#define wEffects wReserved
#endif

BOOL DefaultDCB(DCB* pDcb);

CSerialCom::CSerialCom(void)
{
	m_dwTimeOut = 3000;
	m_hComm = INVALID_HANDLE_VALUE;	
	m_bUseModem = 0;
}

CSerialCom::~CSerialCom(void)
{
	if(m_hComm != INVALID_HANDLE_VALUE)
	{
		CloseHandle(m_hComm);
		m_hComm = INVALID_HANDLE_VALUE;
	}
}


/****************************************************************************
*   Name
		OpenCom
*	Type
		public
*	Function
		Open a communication port. Always open a serial com.
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		pcc
			[in] ComDevice constuct
*****************************************************************************/
BOOL CSerialCom::OpenCom(const ComDevice& pcc)
{
	CDebug::ShowFunMessage(_T("CSerialCom::OpenCom"));

	//serial com name
	TCHAR szPort[32]; 
	//stop bites
	BYTE stopBits;
	//check the baud
	
	if((pcc.baudRate < 110) || (pcc.baudRate) > 256000)
	{ 		   
		CDebug::ShowErroMessage(_T("Buad overflow!"));			
		return FALSE;
	} 
	
	//time out setting
	m_dwTimeOut = pcc.timeOut;
	//the format for the multi-com card
#if defined(_WIN32_WCE_CEPC) || defined(_ARM_)
	_tcscpy(szPort, pcc.deviceName);
	_tcscat(szPort,_T(":"));
#else
	_tcscpy(szPort, _T("\\\\.\\"));	
	_tcscpy(szPort + 4, pcc.deviceName);
#endif
	//open serial com
	HANDLE hComm = CreateFile(szPort, GENERIC_READ | GENERIC_WRITE, 0, 
		NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	//not success, return error message

	CString strComPortMes;
	if(INVALID_HANDLE_VALUE == hComm)
	{
		strComPortMes.Format(_T(" Open %s: Failed"), pcc.deviceName);
		CDebug::ShowDetailMessage((LPTSTR)(LPCTSTR)strComPortMes);
		return FALSE;
	}
	
	//stop bits
	switch(pcc.stopBits)
	{
	case 1:
		stopBits = 0;
		break;
	case 2:
		stopBits = 2;
		break;
	default:
		stopBits = 0;
	}
	//sett DCB 
    DCB dcb;

	//Initial DCB
	DefaultDCB(&dcb);

	//get the default setting
	GetCommState(hComm, &dcb);

	//modify the setting parameters
	dcb.BaudRate = (DWORD)pcc.baudRate;
	dcb.ByteSize = (BYTE)pcc.dataBits;
	dcb.StopBits = stopBits;
	dcb.Parity = (BYTE)pcc.parity;

	/*Record Com Message*/
	TCHAR cCheck;
	switch(pcc.parity)
	{
	case 0:
		cCheck = _T('N');
		break;
	case 1:
		cCheck = _T('O');
		break;
	case 2:
		cCheck = _T('E');
		break;
	case 3:
		cCheck = _T('M');
		break;
	case 4:
		cCheck = _T('S');
		break;
	default:
		ATLASSERT(FALSE);
	}
	strComPortMes.Format(_T(" Open %s: %d,%c,%d,%d "), 
		pcc.deviceName, pcc.baudRate, cCheck, pcc.dataBits,pcc.stopBits);

	CDebug::ShowDetailMessage((LPTSTR)(LPCTSTR)strComPortMes);

	if(!SetCommState( hComm, &dcb ))
	{	
		//not success, close the handle
		CloseHandle(hComm);
		return FALSE;
	}
	
	//setup the serial com's buffer sizes:RX_QUEU = 1024 TX_QUEU = 512
	SetupComm(hComm, RX_QUEU, TX_QUEU);

	//set communication mask
    SetCommMask(hComm, EV_RXFLAG | EV_RXCHAR | EV_BREAK | EV_TXEMPTY);
	// clear the com buffer
	PurgeComm(hComm, PURGE_TXABORT | PURGE_RXABORT | PURGE_TXCLEAR | PURGE_RXCLEAR);	
	//set timeout construct
	
	COMMTIMEOUTS CommTimeOuts;
	//calculate method: ReadTotalTimeout = (ReadTotalTimeoutMultiplier * bytes_to_read)+ ReadToTaltimeoutConstant
	
	CommTimeOuts.ReadIntervalTimeout = MAXDWORD;
	
	CommTimeOuts.ReadTotalTimeoutMultiplier = 0;
	
	CommTimeOuts.ReadTotalTimeoutConstant = 0;
	//calculate mothod £ºWriteTotalTimeout = (WriteTotalTimeoutMultiplier * bytes_to_write)+ WriteToTaltimeoutConstant
	
	CommTimeOuts.WriteTotalTimeoutMultiplier = 2 * CBR_9600 / dcb.BaudRate;
	
	CommTimeOuts.WriteTotalTimeoutConstant = 25;
	
	//set
	SetCommTimeouts(hComm, &CommTimeOuts);
	//store handle
	m_hComm = hComm;	
	m_pcc = pcc;
		return  TRUE;
}

/****************************************************************************
*   Name
		CloseCom
*	Type
		public
*	Function
		Close communication Device.
*	Return Value
		always true.
*	Parameters
		null
*****************************************************************************/
BOOL CSerialCom::CloseCom()
{
	CDebug::ShowFunMessage(_T("CSerialCom::CloseCom"));

	
if (INVALID_HANDLE_VALUE != m_hComm)
	{		
		CloseHandle( m_hComm );
		m_hComm = INVALID_HANDLE_VALUE;
	}
	return TRUE;
}


/****************************************************************************
*   Name
		PhysicalSend
*	Type
		public
*	Function
		send a frame to serial come.
*	Return Value
		return true if succes, else return false.
*	Parameters
	pbyBuf
		[in] send message buffer
	dwWriteLen
		[in] the length to send
*****************************************************************************/
BOOL CSerialCom::PhysicalSend(BYTE* pbyBuf, DWORD dwWriteLen)
{
	CDebug::ShowFunMessage(_T("CSerialCom::PhysicalSend"));
	ASSERT(m_hComm!=INVALID_HANDLE_VALUE) ;
	DWORD dwWrittenLen = 0;
	if(!WriteFile(m_hComm, pbyBuf, dwWriteLen , &dwWrittenLen , NULL))
	{
		CDebug::ShowErroMessage(_T("Send to com failed!"));
		return FALSE;
	}
	
	CDebug::ShowSendMsg(pbyBuf, dwWriteLen);	
	return TRUE;
}

/****************************************************************************
*   Name
		PhysicalReceive
*	Type
		public
*	Function
		receive a frame from serial com.
*	Return Value
		return true if succes, else return false.
*	Parameters
		pbyBuf
			[in] receive message buffer
		iBytesRead
			[in] the length to received
*****************************************************************************/
BOOL CSerialCom::PhysicalReceive(BYTE* pbyBuf, int nStart, int& nBytesRead)
{
	CDebug::ShowFunMessage(_T("CSerialCom::PhysicalReceive"));
	//error status word
	DWORD dwError;
	//current status
	COMSTAT	cs;
	//the number received 
	DWORD dwByteRead = 0;
	//temporary buffer
	BYTE byTemp[512];
	//the expect received length
	int nLen = 0;
	//clear the buffer
	ZeroMemory(byTemp, sizeof(byTemp));
	//get the current serial status
	ClearCommError(m_hComm, &dwError, &cs );		
	//if no message
	if(cs.cbInQue)
	{
		//receive the max length :512 bytes
		if(cs.cbInQue >= 512)
		{
			nLen = 512;
		}
		else
		{
			nLen = cs.cbInQue;
		}
		//read the data for serial com
		BOOL bReadSuc = ReadFile(m_hComm, byTemp, nLen, &dwByteRead, NULL);
		//if success
		if(bReadSuc && dwByteRead > 0)
		{
			CDebug::ShowRecMsg(byTemp, dwByteRead) ;
			//
			int nRemainSpace = MAX_BUFFER - nStart;
			if ( nRemainSpace < dwByteRead) 
			{
				CDebug::ShowErroMessage(_T("Receive buffer overflow, can not to store!"));
				return FALSE;
			}
			else
			{
				CopyMemory( pbyBuf + nStart, byTemp, dwByteRead);
				nBytesRead = dwByteRead;
				return TRUE;
			}
		}
		//if failed 
		else
		{			
			CDebug::ShowErroMessage(_T("Not reveive data!"));
			return FALSE;
		}
	}
	//return the error message
	CDebug::ShowErroMessage(_T("The com not have data!"));
	return FALSE;
}

/****************************************************************************
*   Name
		ReOpenCom
*	Type
		public
*	Function
		open serial com over again.
*	Return Value
		return true if succes, else return false.
*	Parameters
		null
*****************************************************************************/
BOOL CSerialCom::ReOpenCom()
{
	CDebug::ShowFunMessage(_T("CSerialCom::ReOpenCom"));
	if(m_bUseModem)//
	{
		return FALSE;
	}
	CloseCom();
	Sleep(10) ;
	return OpenCom(m_pcc) ;
}

/****************************************************************************
*   Name
DefaultDCB
*	Type
Global
*	Function
Initial the Com Parms
*	Return Value
return true if succes, else return false.
*	Parameters
pDcb
*****************************************************************************/
BOOL DefaultDCB(DCB* pDcb)
{
	if(pDcb == NULL)
	{
		return FALSE;
	}
	pDcb->DCBlength = sizeof(DCB);			// sizeof(DCB) 
	pDcb->BaudRate = CBR_1200;				// current baud rate 
	pDcb->fBinary = TRUE;					// binary mode, no EOF check 
	pDcb->fParity = FALSE;					// enable parity checking 
	pDcb->fOutxCtsFlow = FALSE;				// CTS output flow control 
	pDcb->fOutxDsrFlow = FALSE;				// DSR output flow control 
	pDcb->fDtrControl =  DTR_CONTROL_ENABLE;// DTR flow control type 
	pDcb->fDsrSensitivity = FALSE;			// DSR sensitivity 
	pDcb->fTXContinueOnXoff = FALSE;		// XOFF continues Tx 
	pDcb->fOutX = FALSE;					// XON/XOFF out flow control 
	pDcb->fInX = FALSE;						// XON/XOFF in flow control 
	pDcb->fErrorChar = FALSE;				// enable error replacement 
	pDcb->fNull = FALSE;					// enable null stripping 
	pDcb->fRtsControl = RTS_CONTROL_ENABLE;	// RTS flow control 
	pDcb->fAbortOnError = FALSE;			// abort on error 
	pDcb->fDummy2 = 0;						// reserved 
	pDcb->wReserved = 0;					// not currently used 
	pDcb->XonLim = 2048;					// transmit XON threshold 
	pDcb->XoffLim = 512;					// transmit XOFF threshold 
	pDcb->ByteSize = 7;						// number of bits/byte, 4-8 
	pDcb->Parity = NOPARITY;				// 0-4=no,odd,even,mark,space 
	pDcb->StopBits = ONESTOPBIT;			// 0,1,2 = 1, 1.5, 2 
	pDcb->XonChar = 17;						// Tx and Rx XON character 
	pDcb->XoffChar = 19;					// Tx and Rx XOFF character 
	pDcb->ErrorChar = 0;					// error replacement character 
	pDcb->EofChar = 0;						// end of input character 
	pDcb->EvtChar = 0;						// received event character 
	pDcb->wReserved1 = 0;					// reserved; do not use 
	return TRUE;
}