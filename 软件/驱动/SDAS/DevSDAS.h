/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   DevSDAS.h
Class Name:  CDevSDAS
Brief:       Device sub class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#pragma once

class CDevSDAS: public CDevBase
{
	DECLARE_DYNCREATE(CDevSDAS)
public:
	CDevSDAS(void);
	~CDevSDAS(void);
	
	virtual BOOL GetRegisters(LPVOID *ppReg, int *pRegNum ); 
	virtual BOOL GetDevices(LPVOID *ppDevice, int *pDeviceNum);
	virtual WORD ConvertUserConfigToVar( LPVOID lpDbItemItem, LPVOID lpVar);
	virtual BOOL AddVarToPacket( LPVOID lpVar, int nVarAccessType, LPVOID lpPacket);
	virtual BOOL ProcessPacket2( LPVOID lpPacket );
	virtual BOOL TryConnect( const TCHAR*  szDeviceName, int nUnitAddr, LPVOID lpDevAddr );
		
private:
	int GetSendTimes(PPACKET  pPac);
	BOOL GetSendString(PPACKET pPac, int& nLen, int& nRecLen, int nTimes); 
	BOOL SendDataToKingView(PPACKET pPac, int nLen); 
	BOOL TryConnectOfSelfDefined(const TCHAR*  szDeviceName, int nUnitAddr, PDEVADDR pDevAddr); 
	BOOL GetTryConnectPacketParam(PPACKET pPacket, IdNo2* pIdNo, const TCHAR* szDeviceName);
	BOOL AssertPacketProperty(PPACKET pPac);
	BOOL Transmission(int nLen, int nExpectedLen, int& nRecLen); 
	BOOL PreProcessData (PPACKET pPac, int nbyteRead,int nExpectedLen, int nTimes);
	BOOL GetFirstFrame(int nbyteRead, int nExpectedLen, BYTE byFrameHead);
   

	BOOL GetFrame(BYTE* pByFrame,BYTE** pStart,BYTE** pEnd,int nSearch);
	BOOL CheckFrame(BYTE* pStart,BYTE* pEnd);
	BOOL SetData(BYTE* pStart,BYTE* pEnd);
	BOOL SetData(CString csContent,int id);

	FLOAT m_flData[14];
	int mIndex;
};
