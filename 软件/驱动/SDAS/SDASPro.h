/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   SDASPro.h
Class Name:  CSDASPro
Brief:       Driver project class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#pragma once

#include "icompro.h"

#include "serialcom.h"
/////////////////////////////////////////////////////////////////////////////
// CSDASPro command target

class CSDASPro : public CCmdTarget
{
	DECLARE_DYNCREATE(CSDASPro)		
	CSDASPro();           // protected constructor used by dynamic creation
	
	// Attributes
public:

	
protected:
	
	// Operations
public:
	
	// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSDASPro)
public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL
	
	// Implementation
protected:
	virtual ~CSDASPro();
	
	// Generated message map functions
	//{{AFX_MSG(CSDASPro)
	// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	
	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CSDASPro)		
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSDASPro)
	// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
		
	BEGIN_INTERFACE_PART(ProtocolImp, IProtocolImp)
	STDMETHOD_(BOOL,GetRegisters)(const TCHAR* szDeviceName,LPVOID *ppRegs, int *pRegNum);
	STDMETHOD_(BOOL,TryConnect)( const TCHAR*  pDeviceName, int nUnitAddr, LPVOID lpDevAddr );
	STDMETHOD_(WORD,ConvertUserConfigToVar)(LPVOID lpDbItem, LPVOID lpVar);
	STDMETHOD_(int, LoadDeviceInfo)(const TCHAR* sProd, const TCHAR* sDevName, int nType);
	STDMETHOD_(BOOL, OpenComDevice)( int nDeviceType, LPVOID lpInitData );
	STDMETHOD_(BOOL, CloseComDevice)();
	STDMETHOD_(BOOL, InitialDevice)(const TCHAR* pDeviceName, int nUnitAddr, LPVOID lpDevAddr );
	STDMETHOD_(BOOL,AddVarToPacket)(LPVOID lpVar, int nVarAccessType, LPVOID lpPacket);
		
	STDMETHOD_(int, ProcessPacket)(LPVOID lpPacket);
	STDMETHOD_(BOOL,StrToDevAddr) (const TCHAR* str, LPVOID lpDevAddr );
	STDMETHOD_(TCHAR * , GetLastError)(void);
	
	END_INTERFACE_PART(ProtocolImp)
	BEGIN_INTERFACE_PART(ProtocolImp2,IProtocolImp2)
	STDMETHOD_(int, ProcessPacket2)(LPVOID lpPacket);
	STDMETHOD_(BOOL,SetTrans)( LPVOID* lpHcomm );
	STDMETHOD_(BOOL,SetInitialString)(TCHAR*  pDeviceName, LPVOID lpDevAddr,LPVOID InitialString);
	STDMETHOD_(BOOL,GetTrans)( LPVOID* lpHcomm );
	END_INTERFACE_PART(ProtocolImp2)
		
private:
	
	
public:
	//Attributes
	CSerialCom m_ComObj; //serial com communication class object
	
	int m_nLastErrorCode; //the latest error code
	//Operations
	CDevBase* GetDevObj(const CString& szKind); 
	CDevBase* GetFirstDevObj(); 

private:	
	CPtrList m_DevList ;  //varlist to save device calss objects 
};

/////////////////////////////////////////////////////////////////////////////
//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.



