/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Devbase.cpp
Class Name:  CDevbase
Brief:       Device base class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#include "StdAfx.h"
#include "Devbase.h"
#include "SDASPro.h"

IMPLEMENT_DYNAMIC(CDevBase,CObject)

CDevBase::CDevBase(void)
{
	m_szDevKind = "" ;
	m_dwTimeOut = 3000;
	m_pSerialCom = NULL;

}

CDevBase::~CDevBase(void)
{
}



/****************************************************************************
*   Name
		SetProPtr
*	Type
		public
*	Function
		set the project class pointer 
*	Return Value
		null
*	Parameters
		pPro
			[in] project class pointer.
*****************************************************************************/
void CDevBase::SetProPtr(CSDASPro* pPro)
{	
	m_pPro = pPro;
}

/****************************************************************************
*   Name
		StrToDevAddr
*	Type
		public
*	Function
		Check the input device address string and transform it into lpDevAddr. 
*	Return Value
		return true and transform str into lpDevAddr if successful; otherwise false.
*	Parameters
		str
			[in] Pointer a null-terminated string that identifies address string.
		lpDevAddr
			[in,out] Pointer a DEVADDR variant, see "DATATYPE.H" in detail.
*****************************************************************************/
BOOL CDevBase::StrToDevAddr(const TCHAR* pStr, LPVOID lpDevAddr)
{
	CDebug::ShowFunMessage(_T("CDevBase::StrToDevAddr"));
	
	//Check the parameter pointer
	ASSERT(pStr != NULL);
	ASSERT(lpDevAddr != NULL);
	DEVADDR *pDevAddr = (DEVADDR*)lpDevAddr;
	
	/*
	Notes:
	This function is to Check the input device address string and transform it into lpDevAddr
	we advise you not to modify for common drivers.
	*/
		//initialization register characters 
	CString szAddrStr(pStr);
	
	if(szAddrStr.Find(_T("/IMP_MSG")) != -1)
	{
		CDebug::bShowIMPMsg = 1;
		szAddrStr.Replace(_T("/IMP_MSG"), NULL);
	}
	
	if(szAddrStr.Find(_T("/ERR_MSG")) != -1)
	{
		CDebug::bShowERRMsg = 1;
		szAddrStr.Replace(_T("/ERR_MSG"), NULL);
	}
	
	if(szAddrStr.Find(_T("/FUN_MSG")) != -1)
	{
		CDebug::bShowFUNMsg = 1;
		szAddrStr.Replace(_T("/FUN_MSG"), NULL);
	}
	
	if(szAddrStr.Find(_T("/DATA_MSG")) != -1)
	{
		CDebug::bShowDATAMsg = 1;
		szAddrStr.Replace(_T("/DATA_MSG"), NULL);
	}
	
	if(szAddrStr.Find(_T("/USER_MSG")) != -1)
	{
		CDebug::bShowUSERMsg = 1;
		szAddrStr.Replace(_T("/USER_MSG"), NULL);
	}
	
	if(szAddrStr.Find(_T("/ALL_MSG")) != -1)
	{
		CDebug::bShowALLMsg = 1;
		szAddrStr.Replace(_T("/ALL_MSG"), NULL);
	}

	//Delete blank characters 
	szAddrStr.TrimLeft();
	szAddrStr.TrimRight();
	//
	if (szAddrStr.IsEmpty())
	{
		m_pPro->m_nLastErrorCode = ERR_ADDR_FORMAT;//dcm modify 2007.01.26, same with the following
		return FALSE;
	}
	//Not to access the Maximum length supported. 
	if (szAddrStr.GetLength() > MAX_DEV_LEN)
	{
		m_pPro->m_nLastErrorCode = ERR_ADDR_FORMAT;
		return FALSE;
	}
	
	// Allowed  characters :0123456789ABCDEF.:
	//Can modified according to actual. 
	CString szAllow(_T("0123456789ABCDEF.:"));
	//Check all characters.
	for (int nIndex = 0; nIndex < szAddrStr.GetLength(); nIndex++)
	{
		if(szAllow.Find(szAddrStr[nIndex]) == -1)
		{
			m_pPro->m_nLastErrorCode = ERR_ADDR_FORMAT;
			return FALSE;
		}
	}
	


	//converte to addresses number. 
	UINT nAddress = _ttoi(szAddrStr);
	
	if (nAddress < MIN_SUBSTATION_NUM)
	{
		//if over scope , return the error message.
		m_pPro->m_nLastErrorCode = ERR_ADDR_EXCEED;
		return FALSE;
	}	
	else if(nAddress > MAX_SUBSTATION_NUM)
	{
		//if over scope , return the error message.
		m_pPro->m_nLastErrorCode = ERR_ADDR_EXCEED;
		return FALSE;
	}
	
	pDevAddr->nDevAddr = nAddress;
	
	//save the original string, may be use
	_tcscpy(pDevAddr->sDevAddr, szAddrStr);
	return TRUE;
}

/****************************************************************************
*   Name
		ConvertUserConfigToVar
*	Type
		public
*	Function
		Convert the variant string to KINGVIEW's structural variables (PLCVAR, see "DATATYPE.h"). 
*	Return Value
		return s_ok if successful; otherwise s_false.
*	Parameters
		lpDbItem
			[in]  Pointer to a MiniDbItem variant
			[in,out] Pointer to a PLCVAR variant
*****************************************************************************/
WORD CDevBase::ConvertUserConfigToVar( LPVOID lpDbItemItem, LPVOID lpVar)
{
	
	ASSERT(lpDbItemItem != NULL);
	ASSERT(lpVar != NULL);
	
	MiniDbItem * pDbItem = (MiniDbItem*)lpDbItemItem;
	PPLCVAR pPlcVar = (PPLCVAR) lpVar;
	
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
		m_pPro->m_nLastErrorCode = ERR_REG_NAME; //dcm modify 2007.01.26, same with the following
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
		//
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
		SendDataToIdNo
*	Type
		public
*	Function
		Send Data To IdNo
*	Return Value
		return s_ok if successful; otherwise s_false.
*	Parameters
		pIdNo
			[in,out] Var construct, see "DATATYPE.h" in detail
		buf
			[in,out] a pointer to the  buffer
		iLen
			[out] the received length 
*****************************************************************************/
BOOL CDevBase::SendDataToIdNo(IdNo2* pIdNo, BYTE* pBuf, int nLen)  //
{
	CDebug::ShowFunMessage(_T("CDevBase::SendDataToIdNo"));
	ASSERT(pIdNo != NULL);	
	
	switch (pIdNo->wDataType)
	{
		//bit
	case BIT_DATATYPE:
		pIdNo->plcValue.bitVal = *((BOOL *)pBuf);
		break;
		//byte
	case BYTE_DATATYPE:
		pIdNo->plcValue.byteVal = (BYTE)(*((_TUCHAR *)pBuf));
		break;
		//int
	case INT_DATATYPE:
		pIdNo->plcValue.intVal = *((short * )pBuf);
		break;
		//unsigned int
	case UINT_DATATYPE:
		pIdNo->plcValue.uintVal = *((WORD * )pBuf);
		break;
		//BCD
	case BCD_DATATYPE:
		//		pIdNo->plcValue.bcdVal = WordToBcd(*((WORD * )buf));
		break;
		//long int
	case LONG_DATATYPE:
		pIdNo->plcValue.intVal = *((int * )pBuf);
		break;
		//float
	case FLOAT_DATATYPE:
		pIdNo->plcValue.floatVal = *((float * )pBuf);
		break;
		//long BCD
	case LONGBCD_DATATYPE:
		pIdNo->plcValue.longbcdVal = CHelper::LongToBcd(*((DWORD * )pBuf));			
		break;
		//string
	case STRING_DATATYPE:
		CHelper::KvGlobalGetAtomName(pIdNo->plcValue.strVal, (TCHAR *)pBuf, 255); //not support in EVC
		break;
		//not to arrive here
	default:
		ASSERT(FALSE);
		return FALSE;			
	}
	//get the current time 
			
	CHelper::KvCoFileTimeNow(&pIdNo->ftTimeStamps); //	
	pIdNo->wQualities = COM_QUALITY_GOOD;
	
	return TRUE;
}

