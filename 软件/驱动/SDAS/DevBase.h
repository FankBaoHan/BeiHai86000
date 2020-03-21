/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   DevBase.h
Class Name:  CDevBase
Brief:       Device base class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#pragma once
#include "Reg.h"
#include "Check.h"

class CSDASPro;
typedef CArray<float,float&> CFloatArray ;

class CDevBase: public CObject
{
	DECLARE_DYNCREATE(CDevBase)
		
public:
	CDevBase();
	virtual ~CDevBase(void);
	
	//interface function
	virtual BOOL GetRegisters(LPVOID *ppReg, int *pRegNum )=0; 
	virtual BOOL GetDevices(LPVOID *ppDevice, int *pDeviceNum) =0; 
    virtual WORD ConvertUserConfigToVar( LPVOID lpDbItemItem, LPVOID lpVar);                  //这个不是纯虚的
	virtual BOOL AddVarToPacket( LPVOID lpVar, int nVarAccessType, LPVOID lpPacket)=0;
	virtual BOOL ProcessPacket2( LPVOID lpPacket )=0;
	virtual BOOL TryConnect( const TCHAR*  szDeviceName, int nUnitAddr, LPVOID lpDevAddr ) =0;

	void SetProPtr(CSDASPro* pPro);
	BOOL StrToDevAddr(const TCHAR* pStr, LPVOID lpDevAddr);
	BOOL SendDataToIdNo(IdNo2* pIdNo, BYTE* pBuf, int nLen);
   
public:
	CString	m_szDevKind;                           //device var
	DWORD	m_dwTimeOut;                           //timeout var
	
	CSerialCom	*m_pSerialCom;                     //pointer to serial com class
	CCheck  	*m_pCheck;                         //pointer to check class
	CSDASPro* m_pPro;	           //pointer to project class
	
protected:
	BYTE m_bySndBuf[MAX_BUFFER];	               //send buffer
	BYTE m_byRecBuf[MAX_BUFFER];	               //receive buffer
	BYTE m_byPreRecBuf[MAX_BUFFER];                //last receive buffer
	int m_nPreRecLen;                              //the last remaild data's length
	
};
