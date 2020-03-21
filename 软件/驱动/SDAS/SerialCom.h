/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   SerialCom.h
Class Name:  CSerialCom
Brief:       Serial Communication class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#pragma once
#include "Reg.h"

class CSerialCom
{
public:
	CSerialCom(void);
	~CSerialCom(void);
	//Operations
	
	BOOL OpenCom(const ComDevice& InitData) ; 
	BOOL CloseCom(); //close the serial com
	BOOL ReOpenCom(); //re-open the serial com
	BOOL PhysicalSend(BYTE* pbyBuf,DWORD dwWriteLen); 
	BOOL PhysicalReceive(BYTE* pbyBuf,int nStart, int& nBytesRead);
		
		
public:
	//Attributes
	DWORD m_dwTimeOut; 
	HANDLE m_hComm;
	bool m_bUseModem;
	

private:
	ComDevice m_pcc;
};


