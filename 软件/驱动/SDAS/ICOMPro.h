
#ifndef  __ICOMPRO2_H__
#define  __ICOMPRO2_H__

//The interface ID for this special interface.
#define STR_IID_IPROTOCOL_IMP2  0x8cb34601, 0xa8f5, 0x11d4, { 0x8b, 0xe0, 0x0, 0x50, 0xba, 0xcf, 0xbb, 0x99 } 
static const IID IID_ProtocolImp2 = { STR_IID_IPROTOCOL_IMP2 };

#define STR_IID_IPROTOCOL_IMP  0x2d921821, 0xd190, 0x11d1, { 0xbb, 0x16, 0x0, 0x40, 0x5, 0x62, 0x5a, 0xd7 }
static const IID IID_ProtocolImp = { STR_IID_IPROTOCOL_IMP };

//////////////////////////////////////////////////////////////////////////////
// CProtocolImp: Interface of Protocol.
class IProtocolImp2 : public IUnknown
{
public:	
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** ppvObj) = 0;
    virtual ULONG STDMETHODCALLTYPE AddRef() = 0;
    virtual ULONG STDMETHODCALLTYPE Release() = 0;
	
	//This Interface
	virtual int   STDMETHODCALLTYPE ProcessPacket2( LPVOID lpPacket ) = 0;	
	virtual BOOL   STDMETHODCALLTYPE SetTrans( LPVOID *lpHcomm ) = 0;
	virtual BOOL   STDMETHODCALLTYPE SetInitialString(TCHAR*  pDeviceName, LPVOID lpDevAddr,LPVOID InitialString) = 0;
	virtual BOOL   STDMETHODCALLTYPE GetTrans( LPVOID *lpHcomm ) = 0;
};
class IProtocolImp : public IUnknown
{
public:	
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** ppvObj) = 0;
    virtual ULONG STDMETHODCALLTYPE AddRef() = 0;
    virtual ULONG STDMETHODCALLTYPE Release() = 0;
	
	//This Interface
	virtual BOOL  STDMETHODCALLTYPE GetRegisters( const TCHAR* szDeviceName, LPVOID *ppReg, int *pRegNum ) = 0;
	
	virtual BOOL  STDMETHODCALLTYPE TryConnect( const TCHAR*  pDeviceName, int nUnitAddr, LPVOID lpDevAddr ) = 0;
	
	virtual WORD  STDMETHODCALLTYPE ConvertUserConfigToVar( LPVOID  lpDbItem, LPVOID lpVar ) = 0;

	virtual int   STDMETHODCALLTYPE LoadDeviceInfo( const TCHAR* sProducer, const TCHAR *sDeviceName, int nSubType ) = 0;
	
	
	virtual BOOL  STDMETHODCALLTYPE OpenComDevice( int nDeviceType, LPVOID lpInitData ) = 0;
	
	virtual BOOL  STDMETHODCALLTYPE CloseComDevice() = 0;
	
	virtual BOOL  STDMETHODCALLTYPE InitialDevice(const TCHAR* pDeviceName, int nUnitAddr, LPVOID lpDevAddr ) = 0;
	
	virtual BOOL  STDMETHODCALLTYPE AddVarToPacket( LPVOID lpVar, int nVarAccessType, LPVOID lpPacket) = 0;
	
	virtual int   STDMETHODCALLTYPE ProcessPacket( LPVOID lpPacket ) = 0;
	virtual BOOL  STDMETHODCALLTYPE StrToDevAddr( const TCHAR*  str, LPVOID lpDevAddr )=0;
	virtual TCHAR * STDMETHODCALLTYPE GetLastError()=0;// int * errCode, char ** errStr );
};

#endif    //__PROTOCOLX_H__
