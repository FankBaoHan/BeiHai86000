/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   SDASPro.cpp
Class Name:  CSDASPro
Brief:       Driver project class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/

#include "stdafx.h"
#include "SDAS.h"
#include "SDASPro.h"
#include "reg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSDASPro
IMPLEMENT_DYNCREATE(CSDASPro, CCmdTarget)

CSDASPro::CSDASPro() 
{
	EnableAutomation();	
	CDevBase *pMod = NULL ;
	m_nLastErrorCode = 0;
	
	pMod = new CDevSDAS ;
	pMod->m_pSerialCom = &m_ComObj;
	pMod->SetProPtr(this);
	m_DevList.AddTail(pMod);
	CDebug::CDebug(); 
//	_CrtSetDbgFlag ( _CRTDBG_ALLOC_MEM_DF| _CRTDBG_LEAK_CHECK_DF); //Test memory leak, EVC not support.
	AfxOleLockApp();
	
}

CSDASPro::~CSDASPro()
{
	// To terminate the application when all objects created with
	// 	with OLE automation, the destructor calls AfxOleUnlockApp.
	AfxOleUnlockApp();
	while(!m_DevList.IsEmpty())
	{
		CDevBase * pMod = ( CDevBase* )m_DevList.RemoveTail();  //detele all device objects for varlist.
		delete pMod;
		
	}
}

void CSDASPro::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.	
	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CSDASPro, CCmdTarget)
//{{AFX_MSG_MAP(CTestPro)
// NOTE - the ClassWizard will add and remove mapping macros here.
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSDASPro, CCmdTarget)
//{{AFX_DISPATCH_MAP(CTestPro)
// NOTE - the ClassWizard will add and remove mapping macros here.
//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ITestPro to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {D9202356-742D-4351-9D18-419879B3D82C}
static const IID IID_ITestPro = { 0xD9202356, 0x742D, 0x4351, { 0x9D, 0x18, 0x41, 0x98, 0x79, 0xB3, 0xD8, 0x2C } } ;

BEGIN_INTERFACE_MAP(CSDASPro, CCmdTarget)
INTERFACE_PART(CSDASPro, IID_ITestPro, Dispatch)
INTERFACE_PART(CSDASPro, IID_ProtocolImp, ProtocolImp)
INTERFACE_PART(CSDASPro, IID_ProtocolImp2, ProtocolImp2)
END_INTERFACE_MAP()

//The following is CLSID
#if defined(_UNICODE) || defined(UNICODE)  || defined(_WIN32_WCE_CEPC) || defined(_ARM_) // for UNICODE and EVC compile
// {8108B18D-ACA2-45C4-A850-864B783C95F8}
IMPLEMENT_OLECREATE(CSDASPro, "SDAS.SDASPro", 0x8108b18d, 0xaca2, 0x45c4, 0xa8, 0x50, 0x86, 0x4b, 0x78, 0x3c, 0x95, 0xf8)
#else
// {64D2D5A1-F933-4A59-AC01-02D5B4E911FA} //for MBCS compile
IMPLEMENT_OLECREATE(CSDASPro, "SDAS.SDASPro", 0x64d2d5a1, 0xf933, 0x4a59, 0xac, 0x1, 0x2, 0xd5, 0xb4, 0xe9, 0x11, 0xfa)
#endif 


////////////////////////////////////////////////////////////////////////////
//				ProtocolImp and  ProtocolImp2 IMPLEMENT code
////////////////////////////////////////////////////////////////////////////

//
STDMETHODIMP_(ULONG) CSDASPro::XProtocolImp::AddRef(void)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
    return pThis->ExternalAddRef();
}
//
STDMETHODIMP_(ULONG) CSDASPro::XProtocolImp::Release(void)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
    return pThis->ExternalRelease();
}
//Query the com interface, not to deal 
STDMETHODIMP CSDASPro::XProtocolImp::QueryInterface(REFIID iid, void FAR* FAR* ppvObj)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
    return pThis->ExternalQueryInterface(&iid, ppvObj);
}

STDMETHODIMP_(ULONG) CSDASPro::XProtocolImp2::AddRef(void)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp2);
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CSDASPro::XProtocolImp2::Release(void)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp2);
	return pThis->ExternalRelease();
}

STDMETHODIMP CSDASPro::XProtocolImp2::QueryInterface( REFIID iid, void FAR* FAR* ppvObj)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp2);
	return pThis->ExternalQueryInterface(&iid, ppvObj);
}

////////////////////////////////////////////////////////////////////////////
//	kernel	function below should be called in touchexplorer environment
////////////////////////////////////////////////////////////////////////////

/****************************************************************************
*   Name
		StrToDevAddr
*	Type
		public
*	Function
		Check the input device address string and transforme it into lpDevAddr.
*	Return Value
		return true and transform str into lpDevAddr if successful; otherwise false.
*	Parameters
		str
			[in] Pointer a null-terminated string that identifies address string.
		lpDevAddr
			[in,out] Pointer a DEVADDR variant, see "DATATYPE.H" in detail.
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp::StrToDevAddr(const TCHAR* str, LPVOID lpDevAddr)
{
	METHOD_PROLOGUE(CSDASPro, ProtocolImp); 
	CDebug::ShowImpMessage(_T("CSDASPro::StrToDevAddr"));
	
	//Check the parameter pointer
	ASSERT(str != NULL);
	ASSERT(lpDevAddr != NULL);
	DEVADDR *pDevAddr = (DEVADDR*)lpDevAddr;
	CDevBase* pDev = pThis->GetFirstDevObj() ;	
	if (pDev)
	{
		return pDev->StrToDevAddr(str, pDevAddr) ;
	}
	else
	{
		ASSERT(FALSE) ;
		return FALSE ;
	}
	
	return TRUE;
	
}

/****************************************************************************
*   Name
		GetRegisters
*	Type
		public
*	Function
		The GetRegister retrieves a register list for the specified szDeviceName
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		szDeviceName
			[in] Pointer to a null-terminated string that identifies the register
			list to retrieve
		ppReg
			[out] Pointer to a variable that receives the address of register list
		pRegNum
			[out] Pointer to a variable that receives the size of register list
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp::GetRegisters(const TCHAR* szDeviceName, 
																		   LPVOID * ppRegs, int *pRegNum)
{	
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::GetRegisters"));
    ASSERT(szDeviceName != NULL);
	ASSERT(ppRegs != NULL);
	ASSERT(pRegNum != NULL);
	CDevBase* pDev = pThis->GetDevObj(szDeviceName) ;	
	if (pDev)
	{
		return pDev->GetRegisters(ppRegs,pRegNum) ;
	}
	else
	{
		ASSERT(FALSE) ;
		return FALSE ;
	}
	
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
		lpVar
			[in,out] Pointer to a PLCVAR variant
*****************************************************************************/
STDMETHODIMP_(WORD) CSDASPro::XProtocolImp:: ConvertUserConfigToVar( LPVOID lpDbItemItem, LPVOID lpVar)
{
	METHOD_PROLOGUE(CSDASPro, ProtocolImp); 
	CDebug::ShowImpMessage(_T("CSDASPro::ConvertUserConfigToVar"));
	ASSERT(lpDbItemItem != NULL);
	ASSERT(lpVar != NULL);
	
	MiniDbItem * pDbItem = (MiniDbItem*)lpDbItemItem;
	PPLCVAR pPlcVar = (PPLCVAR) lpVar;
	CDevBase* pDev = pThis->GetDevObj(pDbItem->szDevName);	//
//	CDevBase* pDev = pThis->GetFirstDevObj() ;	
	if (pDev)
	{
		return pDev->ConvertUserConfigToVar(pDbItem, pPlcVar);
	}
	else
	{
		ASSERT(FALSE) ;
		return FALSE ;
	}
	
	return TRUE;
}

/****************************************************************************
*	Brief
		Not used. return true directly.
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp2::SetInitialString(TCHAR* pDeviceName, LPVOID lpDevAddr,LPVOID InitialString)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp2);
	//always return true
	return  TRUE;
}

/****************************************************************************
*   Name
		GetLastError
*	Type
		public
*	Function
		Return to the latest information when  an error operation occurs.
*	Return Value
		Returned a specific null-terminated string;
*	Parameters
		null
*	Remarks
		It Always can be called in TouchExplorer Entironment when an erro occur.
*****************************************************************************/
STDMETHODIMP_(TCHAR *) CSDASPro::XProtocolImp::GetLastError()
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);	
	CDebug::ShowImpMessage(_T("CSDASPro::GetLastError"));
    const int nErrorStringLength = 256;
	static TCHAR error_info[nErrorStringLength] = {0};
	static int nPreviousError = NO_ERROR;  

	//If there are no error
	if(NO_ERROR == pThis->m_nLastErrorCode ) 
	{
		ZeroMemory(error_info, sizeof(error_info));
	}	

	//need to reload resource
	else if( nPreviousError != pThis->m_nLastErrorCode ) 
	{	
		//load the Dynamics  DLL 
		HMODULE  hResource = ::LoadLibrary( _T( "StringResource.dll" ));
		if(hResource)
		{  
			int	nLoaded = LoadString(hResource, pThis->m_nLastErrorCode,error_info, nErrorStringLength);
			FreeLibrary(hResource);

			//not found the error string message.
			if( 0 == nLoaded )
			{
				//_tcscpy( error_info, _T( "Undefined Error;\nPlease reference the handbook!" ) );
				nLoaded = ::LoadString(hResource, ERR_ERRORINFO_NOTFOUND, error_info, nErrorStringLength);
				ASSERT(nLoaded != 0);
				if(nLoaded == 0)
				{	
					//Please down load the new DLL from company's website. 
					_tcscpy(error_info, _T("Resource file (StringResource.dll) is too old,Please update it."));
				}
			}
			else
			{
				nPreviousError = pThis->m_nLastErrorCode;
			}
		}  
		//not found "StringResource.dll"
		else
		{   
			_tcscpy( error_info, _T( "File StringResource.dll missed" )); 
		}
	}

	return error_info;	
}


///////////////////////////////////////////////////////////////////////////////////////////////////
//			Kernel Function below should be called in touchview environment
//////////////////////////////////////////////////////////////////////////////////////////////////

/****************************************************************************
*   Name
		CloseComDevice
*	Type
		public
*	Function
		Open a communication port. Always open a serial com or socket .
*	Return Value
		return true if successful; otherwise false.
*	Parameters
		nDeviceType
			[in] Specifies DeviceType. (No use)
		lpInitData
			[in] Point to an ComDevice variant to initial com
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp::OpenComDevice( int nDeviceType, LPVOID lpInitData)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp); 
	CDebug::ShowImpMessage(_T("CSDASPro::OpenComDevice"));

	ComDevice* pcomdev = static_cast<ComDevice*>(lpInitData);	
	
	return pThis->m_ComObj.OpenCom(*pcomdev);
}


/****************************************************************************
*   Name
		CloseComDevice
*	Type
		public
*	Function
		Close communication Device
*	Return Value
		always true.
*	Parameters
		null 
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp::CloseComDevice()
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::CloseComDevice"));
	return pThis->m_ComObj.CloseCom();
	
}

/****************************************************************************
*   Name
		InitialDevice
*	Type
		public
*	Function
		Initial device for the specified pDeviceName
*	Return Value
		It cannot be called in Virtual COM port project.
*	Parameters
		pDeviceName
			[in] Pointer to a null-terminated string that identifies the 
			specified pDeviceName 
		nUnitAddr
			[in]  Specifies the Unit Address 
		lpDevAddr
			[in]  Point to a DEVADDR variant 
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp::InitialDevice(const TCHAR*  pDeviceName, int nUnitAddr, LPVOID lpDevAddr)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::InitialDevice"));

	return TRUE;	
}


/****************************************************************************
*	Brief
		Not used. return true directly.
*****************************************************************************/
STDMETHODIMP_(int) CSDASPro::XProtocolImp::LoadDeviceInfo( const TCHAR* sProd, const TCHAR * sDevName, int nType )
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::LoadDeviceInfo"));
    return  TRUE;
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
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp:: AddVarToPacket( LPVOID lpVar, int nVarAccessType, LPVOID lpPacket)
{    
	METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::AddVarToPacket"));
	ASSERT(lpVar != NULL);
	ASSERT(lpPacket != NULL);
	ASSERT(nVarAccessType == PT_READ || nVarAccessType == PT_WRITE);
	PPACKET pPac = (PACKET *)lpPacket;
	PPLCVAR pVar = (PLCVAR *)lpVar;	
	CDevBase* pDev = pThis->GetDevObj(pPac->pszDevName);	
	if ( pDev )
	{
		return pDev->AddVarToPacket(lpVar,nVarAccessType,lpPacket) ;
	}
	else
	{
		ASSERT(FALSE) ;
		return FALSE ;
	}	
	return TRUE;
}

/****************************************************************************
*   Name
		ProcessPacket
*	Type
		public
*	Function
		Processing of data packets in 5.0 interface (old interface in brief)
*	Return Value
		return true directly. 
*	Parameters
		lpPacket
			[in,out] Pointer a PACKET variant
*****************************************************************************/
STDMETHODIMP_(int) CSDASPro::XProtocolImp:: ProcessPacket(LPVOID lpPacket)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::ProcessPacket"));
	return TRUE;
}

/****************************************************************************
*   Name
		ProcessPacket2
*	Type
		public
*	Function
		Processing of data packets in 5.0 interface (new interface in brief), must be implemented.
*	Return Value
		return true if successful; otherwise false. 
*	Parameters
		lpPacket
			[in,out] Pointer a PACKET variant
*****************************************************************************/
STDMETHODIMP_(int) CSDASPro::XProtocolImp2:: ProcessPacket2(LPVOID lpPacket)
{
	METHOD_PROLOGUE(CSDASPro, ProtocolImp2);
	CDebug::ShowImpMessage(_T("CSDASPro::ProcessPacket2"));
	//check the pointer
	ASSERT(lpPacket != NULL);
	PPACKET pPack = (PPACKET) lpPacket;
	
	CDevBase* pDev = pThis->GetDevObj(pPack->pszDevName);	
	if (pDev)
	{
		return pDev->ProcessPacket2(lpPacket);
	}
	else
	{
		ASSERT(FALSE);
		return FALSE;
	}
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
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp::TryConnect(const TCHAR*  pDeviceName, 
																		 int nUnitAddr, LPVOID lpDevAddr)
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp);
	CDebug::ShowImpMessage(_T("CSDASPro::TryConnect"));
	ASSERT(pDeviceName != NULL);
	ASSERT(lpDevAddr != NULL);
    CDevBase* pDev = pThis->GetDevObj(pDeviceName);	
	if ( pDev )
	{
		return pDev->TryConnect(pDeviceName,nUnitAddr,lpDevAddr);
	}
	else
	{
		ASSERT(FALSE) ;
		return FALSE ;
	}	
	return FALSE ;
}

/****************************************************************************
*   Name
		SetTrans
*	Type
		public
*	Function
		Set the inner-communication handle.
*	Return Value
		Always return true. 
*	Parameters
		lpHcomm
			[in] Pointer to a handle that Specifies an valid-communication handle.
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp2::SetTrans( LPVOID* pHcomm )
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp2);
	CDebug::ShowImpMessage(_T("CSDASPro::SetTrans"));
	//µÃµ½¾ä±ú
	//get the handle
	pThis->m_ComObj.m_hComm = *pHcomm;
	pThis->m_ComObj.m_bUseModem = true;
	return  TRUE;
}


/****************************************************************************
*   Name
		GetTrans
*	Type
		public
*	Function
		Retrieves an inner-communication handle in serial communication project.
*	Return Value
		Always return true. 
*	Parameters
		lpHcomm
			[out] Pointer to a handle that retrieve inner-communication handle.
*****************************************************************************/
STDMETHODIMP_(BOOL) CSDASPro::XProtocolImp2::GetTrans( LPVOID* pHcomm )
{
    METHOD_PROLOGUE(CSDASPro, ProtocolImp2);	
	CDebug::ShowImpMessage(_T("CSDASPro::GetTrans"));
	ASSERT(pHcomm);	
	//·µ»Ø¾ä±ú
	//return the handle
	*pHcomm  = pThis->m_ComObj.m_hComm;
	return TRUE;
}

/****************************************************************************
*   Name
		GetDevObj
*	Type
		public
*	Function
		Get the device class object pointer from varlist.
*	Return Value
		return the pointer's address if successful; otherwise return NULL. 
*	Parameters
		szKind
			[in] Device type.
*****************************************************************************/
CDevBase* CSDASPro::GetDevObj(const CString& szKind)
{
	DEVICE_INFO * pDeviceInfo = NULL;
	int m_iDeviceNum = 0;
	POSITION pos = m_DevList.GetHeadPosition();	
	
	while (pos)
	{
		CDevBase* pDev = ( CDevBase* )m_DevList.GetNext(pos);
		pDev->GetDevices((LPVOID *)&pDeviceInfo,  &m_iDeviceNum);
		for(int iDeviceIndex = 0; iDeviceIndex < m_iDeviceNum; iDeviceIndex++)		
		{
			if(pDeviceInfo[iDeviceIndex].sDeviceName == szKind)
			{
				return pDev;				
			}			
		}		
	}	
	return NULL;
}

/****************************************************************************
*   Name
		GetFirstDevObj
*	Type
		public
*	Function
		Get the device class object pointer .
*	Return Value
		return the pointer's address if successful; otherwise return NULL. 
*	Parameters
		null
*****************************************************************************/
CDevBase* CSDASPro::GetFirstDevObj()
{
	DEVICE_INFO * pDeviceInfo = NULL;
	int m_iDeviceNum = 0;
	POSITION pos = m_DevList.GetHeadPosition();		
	if (pos)
	{
		CDevBase* pDev = ( CDevBase* )m_DevList.GetNext(pos);
		pDev->GetDevices((LPVOID *)&pDeviceInfo,  &m_iDeviceNum);	
		if(pDeviceInfo[0].sDeviceName != _T(""))
		{
			return pDev;				
		}			
	}	
	return NULL;
}