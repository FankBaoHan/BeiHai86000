/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   DevSDAS.cpp
Class Name:  CDevSDAS
Brief:       Device sub class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#include "StdAfx.h"
#include "DevSDAS.h"
#include "SDASPro.h"

//add or modify devices  here  
static DEVICE_INFO gsDeviceInfo[]=
{
	{_T("SDAS")  }
};

#define		DEVICE_TYPE_NUM		sizeof(gsDeviceInfo)/sizeof(DEVICE_INFO)  

//Device variable define, add your code  here, the following for your reference
#define		DEVICE_NAME1			_T("SDAS")
//#define		DEVICE_NAME2			_T("Name2")



static REG_INFO gsRegInfo[]=
{
	{_T("D"), 0, 13, FLOAT_DATATYPE, PT_READ }
};

#define		REG_TYPE_NUM		sizeof(gsRegInfo)/sizeof(REG_INFO)  

//Register variable define, add your code  here, the following for your reference
#define     D_REG   0

#define		NUM_OF_IDENTIFER 11

IMPLEMENT_DYNAMIC(CDevSDAS, CDevBase)

CDevSDAS::CDevSDAS(void)
{
	m_nPreRecLen = 0; 
}


CDevSDAS::~CDevSDAS(void)
{
}

/****************************************************************************
*   Name
		GetRegisters
*	Type
		public
*	Function
		The GetRegister retrieves a register list for the specified szDeviceName
*	Return Value
		return true if successful;
*	Parameters
		ppReg
			[out] Pointer to a variable that receives the address of register list
		pRegNum
			[out] Pointer to a variable that receives the size of register list
*****************************************************************************/
BOOL CDevSDAS::GetRegisters( LPVOID *ppReg, int *pRegNum )
{
	CDebug::ShowFunMessage(_T("CDevSDAS::GetRegisters"));
	ASSERT(ppReg != NULL);
	ASSERT(pRegNum != NULL);
	*ppReg = gsRegInfo ;
	*pRegNum = REG_TYPE_NUM ;
	return TRUE ;
}

/****************************************************************************
*   Name
		GetDevices
*	Type
		public
*	Function
		The GetRegister retrieves a devices list for the specified device kind class
*	Return Value
		return true if successful; 
*	Parameters
		ppDevice
			[out] Pointer to a variable that receives the address of device list
		pDeviceNum
			[out] Pointer to a variable that receives the size of device list
*****************************************************************************/
BOOL CDevSDAS::GetDevices(LPVOID *ppDevice, int *pDeviceNum) //
{
	CDebug::ShowFunMessage(_T("CDevName1::GetDevices"));
	ASSERT(ppDevice != NULL);
	ASSERT(pDeviceNum != NULL);
	*ppDevice = gsDeviceInfo ;
	*pDeviceNum = DEVICE_TYPE_NUM ;
	return TRUE ;
}

/****************************************************************************
*   Name
		ConvertUserConfigToVar
*	Type
		public
*	Function
		Convert the variable string to KINGVIEW's structural variables (PLCVAR, see "DATATYPE.h").
*	Return Value
		return 0 if successful; otherwise retrun 1 if failed. 
*	Parameters
		lpDbItem
			[in]  Pointer to a MiniDbItem variant
			[in,out] Pointer to a PLCVAR variant
*****************************************************************************/
WORD CDevSDAS::ConvertUserConfigToVar( LPVOID lpDbItemItem, LPVOID lpVar)  //
{
	
	ASSERT(lpDbItemItem != NULL);
	ASSERT(lpVar != NULL);
	
	MiniDbItem * pDbItem = (MiniDbItem*)lpDbItemItem;
	PPLCVAR pPlcVar = (PPLCVAR) lpVar;
	/*
	Notes:
	This function is to convert the variable string to KINGVIEW's structural variables
	we advise you not to modify for common drivers.
	*/

	REG_INFO * pRegInfo = NULL;
	int nRegNum = 0 ;
	GetRegisters( (LPVOID*)&pRegInfo, &nRegNum );
	//Initialization 
	CString szRegister(pDbItem->szRegister);
	//Delete blank characters 
	szRegister.TrimLeft();
	szRegister.TrimRight();
	if (szRegister.IsEmpty())
	{		
		m_pPro->m_nLastErrorCode = ERR_REG_NAME;
		return 1;
	}
	//Not to access the Maximum length supported.
	if (szRegister.GetLength() > MAX_CONFIG_STRING_LENGTH)
	{
		//m_pPro->m_nLastErrorCode = ERR_CONFIG_MAXLEN;
		m_pPro->m_nLastErrorCode = ERR_REG_NAME;
		return 1;
	}	
	
	//if found
	BOOL bMatched = FALSE;
	//get the index number of register
	int  nRegIndex = 0;
	do
	{	
		//if found the register string
		if (szRegister.Find(pRegInfo[nRegIndex].sRegName) == 0 )
		{
			bMatched = TRUE;
		}
		
	}while (!bMatched && ++nRegIndex < nRegNum );
	
	if (!bMatched)
	{
		m_pPro->m_nLastErrorCode = ERR_REG_NAME;
		return 1;
	}	
	
	
	//get channel string
	CString szChannelConfig = szRegister.Right(szRegister.GetLength() \
		- (int)_tcslen(pRegInfo[nRegIndex].sRegName));
	
	//Allowed characters : 0123456789ABCDEF.:
	CString szAllow(_T("0123456789ABCDEF.:"));
	
	//Check all characters.
	for (int nIndex = 0; nIndex < szChannelConfig.GetLength(); nIndex++)
	{
		
		if (szAllow.Find(szChannelConfig[nIndex]) == -1)
		{
			m_pPro->m_nLastErrorCode = ERR_REG_NAME;
			return 1;
		}
	}
	
	
	//check the string format
	if(pRegInfo[nRegIndex].nLowIndex == pRegInfo[nRegIndex].nUpperIndex)
	{		
		if (pRegInfo[nRegIndex].nLowIndex != 0)
		{
			if(szChannelConfig.IsEmpty())
			{
				m_pPro->m_nLastErrorCode = ERR_REG_CH;
				return 1;
			}
		}
	}	

	//get the channel(offset) number
	int nNo = _ttoi(szChannelConfig);

	//if low the index number less than the high index number
	if(pRegInfo[nRegIndex].nLowIndex < pRegInfo[nRegIndex].nUpperIndex)
	{		

		if(nNo < pRegInfo[nRegIndex].nLowIndex) 
		{
			//if not in the range
			m_pPro->m_nLastErrorCode = ERR_REG_CH;
			return 1; 					
		}	
		else if(nNo >pRegInfo[nRegIndex].nUpperIndex)
		{

			m_pPro->m_nLastErrorCode = ERR_REG_CH;
			return 1; 
		}
		else
		{
			//success
			pPlcVar -> nNo = nNo;								
		}			
	}

	else if (pRegInfo[nRegIndex].nLowIndex ==  0 && pRegInfo[nRegIndex].nUpperIndex == 0)
	{
		if(nNo > 0)  
		{
			m_pPro->m_nLastErrorCode = ERR_REG_CH;
			return 1;
		}
		else if(nNo < 0)
		{
			m_pPro->m_nLastErrorCode = ERR_REG_CH;
			return 1;
		}
		//default
		pPlcVar->nNo = 0;		
	}
	else
	{
		m_pPro->m_nLastErrorCode = ERR_REG_FORMAT;
		return 1;
	}
	
	pPlcVar->nRegType = nRegIndex;						
	pPlcVar->pszRegName = pRegInfo[nRegIndex].sRegName;
	
	if(pDbItem->nDataType & pRegInfo[nRegIndex].wDataType)
	{
		pPlcVar->nDataType = pDbItem->nDataType;		
		return 0;
	}			
	else
	{				
		m_pPro->m_nLastErrorCode =  ERR_REG_DATATYPE;
		return 1;				
	}
	return 0;
}

/****************************************************************************
*   Name
		AddVarToPacket
*	Type
		public
*	Function
		Whether or not the var can add to the lpPacket
*	Return Value
		return true if the lpVar can be added into the lpPacket;otherwise false.
*	Parameters
		lpVar 
			[in]  Pointer to a PLCVAR variant whether or not add into the packet
		nVarAccessType
			[in]  Specifies the lpvar Access Type.(NO USED)
		lpPacket
			[int] Pointer to a PACKET variant whether or not contain the lpVar
*****************************************************************************/
BOOL CDevSDAS::AddVarToPacket( LPVOID lpVar, int nVarAccessType, LPVOID lpPacket)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::AddVarToPacket"));
	PPACKET pPac = (PACKET *)lpPacket;
	PPLCVAR pVar = (PLCVAR *)lpVar;
	/*
	Add you code here
	the following code for your reference.
	*/

	//total vars in the packet
	int nTotalNo;
	//only if var type, address, data types are all the same
	if((nVarAccessType == pPac->nPacketType)&&(pPac->nUnitNo == pVar->nUnitNo)
		&& (pPac->nRegType == pVar->nRegType))
	{
		
		if ( nVarAccessType == PT_WRITE)
		{
			if ( pVar->nNo == pPac->nStartNo)  
			{
				return TRUE;
			}
			else  
			{
				return FALSE;
			}
		}
		//read
		//if var channel less than the packet start No
		if ( pVar->nNo < pPac->nStartNo )
		{//chart:		-----pVar->nNo-------pPac->Start================pPac->End----------
			//get the channel rangle
			nTotalNo = pPac->nEndNo - pVar->nNo + 1;
			//MAX_PACKET_NUM: the allowed Max var numbers in a packet
			
			if ( nTotalNo <= MAX_PACKET_NUM )
			{
				pPac->nStartNo = pVar->nNo;	
				return TRUE;
			}
		}
		//if var channel larger than the packet start No
		else if ( pVar->nNo > pPac->nEndNo )
		{//chart:------------pPac->Start================pPac->End----pVar->nNo------
				//get the channel rangle
			nTotalNo = pVar->nNo - pPac->nStartNo + 1;
			
			if ( nTotalNo <= MAX_PACKET_NUM )
			{
				pPac->nEndNo = pVar->nNo;
				return TRUE;
			}
		}
		else
		{//chart:------------pPac->Start====pVar->nNo============pPac->End----------
			return TRUE;
		}
		//
		//chart:------------pPac->Start====pVar->nNo============pPac->End----------
	}	
	return FALSE;
}

/****************************************************************************
*   Name
		ProcessPacket2
*	Type
		public
*	Function
		Processing of data packets in 6.0 interface (new interface in brief), must be implemented.
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		lpPacket
			[in,out] Pointer a PACKET variable.
*****************************************************************************/
BOOL CDevSDAS::ProcessPacket2( LPVOID lpPacket )
{
	CDebug::ShowFunMessage(_T("CDevSDAS::ProcessPacket2"));
	
	ASSERT(lpPacket != NULL);
	PPACKET pPac = (PPACKET)lpPacket;
	ZeroMemory(m_bySndBuf, sizeof(m_bySndBuf));
	ZeroMemory(m_byRecBuf, sizeof(m_byRecBuf));
	if( !AssertPacketProperty(pPac) )
	{
		return FALSE;
	}

    int nSendLen = 0;                                //send length
	int nExpectedLen = 27;                            //expect length to receive
	int nRecLen = 0;                                 //received lenght

	int nSendTimes = GetSendTimes(pPac);
	
	int  nDealedTime = 1;

	do
	{	
		if (!Transmission(nSendLen, nExpectedLen,  nRecLen))
		{
			CDebug::ShowErroMessage(_T("Transmission error!"));
			return FALSE;
		}
		
		//pre-process the received data
		if (!PreProcessData(pPac, nRecLen, nExpectedLen, nDealedTime))
		{
			CDebug::ShowErroMessage(_T("PreProcessData error!"));
			return FALSE;
		}
		nDealedTime++; 
		
	}while (nDealedTime <= nSendTimes);
	//send correct data to KingView
	if (pPac->nPacketType == PT_READ) 
	{
		return SendDataToKingView(pPac,nExpectedLen);
	}
	return TRUE;
}


/****************************************************************************
*   Name
		TryConnect
*	Type
		public
*	Function
		Try to connect with the Device when the communication failed.
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		pDeviceName
			[in] Pointer to a null-terminated string that identifies the 
			specified pDeviceName.
		nUnitAddr
			[in]  Specifies the Unit Address 
		lpDevAddr
			[in]  Point to a DEVADDR variant 
*	Remarks
		When "processpacket" failed, tryconnect twice,
		if failed again, the touchvew will connect with setted time. 
*****************************************************************************/
BOOL CDevSDAS::TryConnect( const TCHAR*  szDeviceName, int nUnitAddr, LPVOID lpDevAddr )
{
	CDebug::ShowFunMessage(_T("CDevSDAS::TryConnect"));
	ASSERT( m_pSerialCom);
	ASSERT(szDeviceName != NULL);
	ASSERT(lpDevAddr != NULL);
	
		
#if defined(_TRYCONNECT_TRUE)
		//not to try connect, return true directly
	return TRUE;	
#elif defined(_TRYCONNECT_USERDEFINE)
	//try connect by youself 
	DEVADDR *pRealDevAddr = (DEVADDR*)lpDevAddr;
#if defined (_REOPENCOM)
	//re-open the serial port
	m_pSerialCom->ReOpenCom(); 
#endif
	return TryConnectOfSelfDefined(szDeviceName, nUnitAddr, pRealDevAddr);	
#elif defined(_TRYCONNECT_USEPACKET)	
	//try connect by process packet

#if defined (_REOPENCOM)
	m_pSerialCom->ReOpenCom(); 
#endif
	PPACKET   pPac = new PACKET;
	ID_NO2    *pTempIdNo = new ID_NO2;	
	pPac->nUnitNo = nUnitAddr;	
	pTempIdNo->wNo = 0;
	//return value
	BOOL bRet = FALSE;
	//get the packet paramters
	bRet = GetTryConnectPacketParam(pPac, pTempIdNo, szDeviceName);
	
	if(bRet)
	{
		//added to tail of the vallist 
		pPac->varList.AddTail(pTempIdNo);
		//excute
		bRet = ProcessPacket2(pPac);
	}
	delete   pPac;
	delete   pTempIdNo;
	return   bRet;
#endif
	return TRUE;
}


//
BOOL CDevSDAS::TryConnectOfSelfDefined(const TCHAR*  szDeviceName, int iUnitAddr, PDEVADDR pDevAddr)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::VTTryConnect"));
	//This function is to realize  try connect by yourself.
	//Add your  handler code here, but a recommend you to use the Transmission function

	return TRUE;	
}
//
BOOL CDevSDAS::GetTryConnectPacketParam(PPACKET pPacket, IdNo2* pIdNo, const TCHAR* szDeviceName)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::VTGetTryConnectPacketParam"));
	 
	//This function is to build your packet.
	//Add your  handler code here, the following code for your reference
	pPacket->nPacketType = PT_READ;
	pPacket->nRegType = D_REG;	
	pPacket->nStartNo = 0;
	pPacket->nEndNo = 13;
	
	pIdNo->wDataType = UINT_DATATYPE;
	return TRUE;
}

/****************************************************************************
*   Name
AssertPacketProperty
*	Type
private
*	Function
check if the packet property of R/W conflict with the register property of R/W
*	Return Value
return true if correct; otherwise wrong.
*	Parameters
pPac
[in] Pointer a PACKET variant
*****************************************************************************/
BOOL CDevSDAS::AssertPacketProperty(PPACKET pPac)
{
	ASSERT(pPac != NULL);
	int nRegIndex = pPac->nRegType;
	SHORT nPacType = pPac->nPacketType;
	if(gsRegInfo[nRegIndex].nData == PT_READWRITE )
	{
		return TRUE;
	}
	else if(gsRegInfo[nRegIndex].nData != nPacType)
	{
		CString szErrMes;
		if(gsRegInfo[nRegIndex].nData == PT_WRITE)
		{
			szErrMes.Format(_T("Read %s Register Wrong : %s Register is WRITE_ONLY"), gsRegInfo[nRegIndex].sRegName, gsRegInfo[nRegIndex].sRegName);
		}
		else
		{
			szErrMes.Format(_T("Write %s Register Wrong : %s Register is READ_ONLY"), gsRegInfo[nRegIndex].sRegName, gsRegInfo[nRegIndex].sRegName);
		}
		CDebug::ShowDetailMessage((LPTSTR)(LPCTSTR)szErrMes);
		return FALSE;
	}
	else
	{
		return TRUE;
	}
}

/****************************************************************************
*   Name
		GetSendString
*	Type
		public
*	Function
		Build sending frame according the packet information.
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		pPac
			[in,out] Pointer a PACKET variant
		iLen
			[out] the data's length to send
		iRecLen
			[out] the expect length to receive
		iTimes
			[in] The number to deal packet in a process
			 used for some devices which get the data through query-ack for a few times 
*	Remarks
		When "processpacket" failed, tryconnect twice,
		if failed again, the touchvew will connect with setted time. 
*****************************************************************************/
BOOL CDevSDAS::GetSendString(PPACKET pPac,int& nLen, int& nRecLen, int nTimes)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::GetSendString"));
	ASSERT(pPac != NULL);

/*
Notes:
This function is to build the send frame according to communication protocol.

Add your  handler code here,

m_bySndBuf: The send buffer,you should copy the send frame to it
nLen:  the data's length to send
nRecLen: the expect length to receive

you can see the demo example for reference.
*/

	return TRUE;
}

/****************************************************************************
*   Name
		SendDataToKingView
*	Type
		public
*	Function
		Send Data To KingView
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		pPac
			[in,out] Pointer a PACKET variant
		iLen
			[out] the received length 
*****************************************************************************/
BOOL CDevSDAS::SendDataToKingView(PPACKET pPac, int nLen)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::SendDataToKingView"));
	ASSERT(pPac != NULL);
	//
	POSITION pos =  pPac->varList.GetHeadPosition();
	
	while(pos)
	{
		ID_NO2 * pIdNo = static_cast<ID_NO2*>(pPac->varList.GetNext(pos));

		if(pIdNo->wNo<0||pIdNo->wNo>13)
		{
			CDebug::ShowFunMessage(_T("SendDataToKingView:wNo overflow!"));
			return FALSE;
		}

		if(pIdNo->wDataType==FLOAT_DATATYPE)
			pIdNo->plcValue.floatVal=m_flData[pIdNo->wNo];

		CHelper::KvCoFileTimeNow(&pIdNo->ftTimeStamps);
		pIdNo->wQualities = COM_QUALITY_GOOD;
	}

	return TRUE;
}


/****************************************************************************
*   Name
		Transmission
*	Type
		public
*	Function
		Send Data to serial port and receive data from serial port
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		iLen
			[in] The length to send
		iExpectedLen
			[in,out] the expected length to receive
		iRecLen
			[in,out] the received length 
*****************************************************************************/
BOOL CDevSDAS::Transmission(int nLen, int nExpectedLen, int& nRecLen)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::Transmission"));	

	//Clear the serial com
	PurgeComm(m_pSerialCom->m_hComm, PURGE_TXABORT | PURGE_RXABORT | PURGE_TXCLEAR | PURGE_RXCLEAR);
	
	int nByteRead = 0;

	nByteRead = nRecLen;
	
	int nLastStatus = ERR_PREPROCESSDATA_SUCCEED; 
	// first offset:0
	//first  ______________________________________offset = 0
	//secod  XXXXXXXXXX____________________________offset = 10
	int nStartOffset = 0;
	
	ULONGLONG  uTimeOut = CHelper::KvGetTickCount() + m_pSerialCom->m_dwTimeOut;

	ZeroMemory(m_byRecBuf, sizeof(m_byRecBuf));
	BOOL bRet = TRUE;
	
	while(TRUE)
	{		
		//receive data from serial port	
		if(m_pSerialCom->PhysicalReceive(m_byRecBuf, nStartOffset, nByteRead)) 
		{   			
			nStartOffset += nByteRead;

			if(*(m_byRecBuf + nStartOffset - 2) == '\r' && *(m_byRecBuf + nStartOffset - 1) == '\n')
			{
				CDebug::ShowFunMessage(_T("CDevSteerGear::Receive the identifier!"));	
				bRet = TRUE;
				break;
			}
			else if(nStartOffset >= nExpectedLen)
			{
				bRet = TRUE;
				break;
			}
			else
			{
				nByteRead = nExpectedLen - nStartOffset; 
				CDebug::ShowErroMessage(_T("The recieve data not access to expect lenth!"));
				ASSERT(nByteRead >= 0);
				continue;
			}			
		}
		else  
		{
			//default: WAIT_INTERVAL = 5
			Sleep(WAIT_INTERVAL);
		}
		
		//if timeout 
		if(CHelper::KvGetTickCount() > uTimeOut) 
		{			
			m_pPro->m_nLastErrorCode = ERR_TRANSSMISSION_TIMOUT;
			CDebug::ShowErroMessage(_T("Receive from com timeout!"));
			bRet = FALSE;
			break;
		}
	}
	nRecLen = nStartOffset;	
	
	return bRet;
	
}

/****************************************************************************
*   Name
		GetFirstFrame
*	Type
		public
*	Function
		Get a first frame from the received data
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		nbyteRead
			[in] The length received
		iExpectedLen
			[in] the expected length to receive
		chFrameHead
			[in] the first byte (the flag) of the frame 
*****************************************************************************/
BOOL CDevSDAS::GetFirstFrame(int nbyteRead, int nExpectedLen, BYTE byFrameHead)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::GetFirstFrame"));
	ASSERT(nbyteRead >= nExpectedLen);
	//Add your modification  handler code here,
	//the following code for reference
	int i = 0;
	do
	{
		if(m_byRecBuf[i] == byFrameHead)
		{
			memcpy(m_byRecBuf, m_byRecBuf+i, nExpectedLen);
			return TRUE;
		}
		i++;
	}while(i < nbyteRead);
	return FALSE;
}

/****************************************************************************
*   Name
		PreProcessData
*	Type
		public
*	Function
		Pre-Process data when received
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		pPac
			[in,out] Pointer a PACKET variant
		nbyteRead
			[in] The length received
		iExpectedLen
			[in] the expected length to receive
*****************************************************************************/
BOOL CDevSDAS::PreProcessData (PPACKET pPac, int nbyteRead ,int nExpectedLen, int nTimes)
{
	CDebug::ShowFunMessage(_T("CDevSDAS::PreProcessData"));
	ASSERT(pPac != NULL);
	if (pPac == NULL)
	{
		return FALSE;
	}

	BYTE* pStart=new BYTE;
	BYTE* pEnd=new BYTE;
	int frameLength(0);

	if(!GetFrame(m_byRecBuf,&pStart,&pEnd,nbyteRead))
	{
		CDebug::ShowErroMessage(_T("Can not find '$' or '*'!"));
		return FALSE;
	}

	if(!CheckFrame(pStart,pEnd))
	{
		CDebug::ShowErroMessage(_T("The checksum is not correct!"));
		return FALSE;
	}

	if(!SetData(pStart,pEnd))
	{
		CDebug::ShowErroMessage(_T("SetData failed!"));
		return FALSE;
	}

	return TRUE;
}


/****************************************************************************
*   Name
		GetSendTimes
*	Type
		public
*	Function
		According the regsiter information,
		set the query-acknowledge times in a data process.
*	Return Value
		a integer
*	Parameters
		pPac
			[in,out] Pointer a PACKET variant
*****************************************************************************/
int CDevSDAS::GetSendTimes(PPACKET  pPac)//
{
	ASSERT(pPac != NULL);
	int nReturnTimes = 1;	
	//Add your  handler code here
	//the default return value is 1

	return nReturnTimes;
	
}

/****************************************************************************
*   Name
		GetFrame
*	Type
		Private
*	Function
		Get the location of '*'and '$'
*	Return Value
		BOOL
*	Parameters
		[in]The data to be analyzed.
		[in,out]The point of the location of the '$'.
		[in,out]The point of the location of the '*'.
		[in]Number of data received.
*****************************************************************************/
BOOL CDevSDAS::GetFrame(BYTE* pByFrame,BYTE** pStart,BYTE** pEnd,int nSearch)
{
	//CDebug::ShowRecMsg(pByFrame,nSearch);

	int nTmp=nSearch;
	while(nSearch-->0)
	{
		if(*(pByFrame++)=='$')
		{
			*pStart=pByFrame-1;
			break;
		}else if(nSearch==0){
			return FALSE;
		}
	}
	
	BYTE* pTmp=*pStart;
	while(nTmp-->0)
	{
		if(*(pTmp++)=='*')
		{
			*pEnd=pTmp-1;
			break;
		}else if(nTmp==0){
			return FALSE;
		}	
	}

	//CDebug::ShowRecMsg(*pStart+1,*pEnd-*pStart-1);

	return TRUE;
}

/****************************************************************************
*   Name
		CheckFrame
*	Type
		Private
*	Function
		Check the checksum
*	Return Value
		BOOL
*	Parameters
		[in]The location of the start of the frame
		[in]The location of the end of the frame
*****************************************************************************/
BOOL CDevSDAS::CheckFrame(BYTE* pStart,BYTE* pEnd)
{
	ASSERT(pStart != NULL);
	ASSERT(pEnd != NULL);

	if((pEnd- pStart- 1)<=0||(pEnd- pStart- 1)>70)
		return FALSE;

	CDebug::ShowRecMsg(pStart+1,pEnd- pStart- 1);

	BYTE checkSum = CCheck::Xor_Check(pStart + 1,pEnd- pStart- 1);
	BYTE checkRec = CHelper::ASCToByte(pEnd+1);

	if (checkSum == checkRec)
		return TRUE;
	else
	{
		CDebug::ShowRecMsg(&checkSum,1);
		CDebug::ShowRecMsg(&checkRec,1);
		return FALSE;
	}
}

/****************************************************************************
*   Name
		SetData
*	Type
		Private
*	Function
		Set the data
*	Return Value
		BOOL
*	Parameters
		[in]The location of the start of the frame
		[in]The location of the end of the frame
*****************************************************************************/
BOOL CDevSDAS::SetData(BYTE* pStart,BYTE* pEnd)
{
	ASSERT(pStart != NULL);
	ASSERT(pEnd != NULL);
	
	short numOfData=pEnd-pStart+1,i=0,j=0,plcDlm[NUM_OF_IDENTIFER];
	while(numOfData-->0)
	{
		if( *(pStart+i)=='$'||
			*(pStart+i)==','||
			*(pStart+i)=='*')
		{
			if(j<NUM_OF_IDENTIFER){
				plcDlm[j]=i;j++;
			}else{
				CDebug::ShowErroMessage(_T("Too many identifiers!"));
				return FALSE;
			}
		}
		i++;
	}

	if(j!=NUM_OF_IDENTIFER)
	{
		CDebug::ShowErroMessage(_T("Too less identifiers!"));
		return FALSE;
	}

	//SetData
	for(i=0;i<NUM_OF_IDENTIFER-1;i++)
	{
		BYTE byTmp[10]="\0";//
		
		int len=plcDlm[i+1]-plcDlm[i]-1;
		if(len>10)
		{
			CDebug::ShowErroMessage(_T("Data is too long![SetData]"));
			return FALSE;
		}

		memcpy( byTmp, pStart+plcDlm[i]+1,len);

		CString csTmp(byTmp);

		if(!SetData(csTmp,i))
		{
			CDebug::ShowErroMessage(_T("SetData default failed!"));
			return FALSE;
		}
	}

	return TRUE;
}

/****************************************************************************
*   Name
		SetData
*	Type
		Private
*	Function
		Set the data
*	Return Value
		BOOL
*	Parameters
		[in]The content of data
		[in]The id of data
*****************************************************************************/
BOOL CDevSDAS::SetData(CString csContent,int id)
{
	
	//one by one
	switch(id)
	{
		case 0:
			if(csContent.Right(5)!="RDFIR")
				return FALSE;
			break;
		case 1:

			break;
		case 2:

			break;
		case 3:

			break;
		case 4:
			this->mIndex = atoi(csContent);
			
			if (mIndex > 7 || mIndex < 0) {
				CDebug::ShowErroMessage(_T("Cargo number exceeded!"));

				return FALSE;
			}

			break;
		case 5:

			break;
		case 6:

			break;
		case 7:
			if (csContent == "A") {
				m_flData[(mIndex - 1) * 2] = 1.0;
				m_flData[(mIndex - 1) * 2 + 1] = 0.0;
			}

			if (csContent == "V") {
				m_flData[(mIndex - 1) * 2] = 0.0;
				m_flData[(mIndex - 1) * 2 + 1] = 0.0;
			}
			
			if (csContent == "X") {
				m_flData[(mIndex - 1) * 2 + 1] = 1.0;
			}

			break;
		case 8:

			break;
		case 9:

			break;
		default:
			return FALSE;
			break;
	}

	return TRUE;
}