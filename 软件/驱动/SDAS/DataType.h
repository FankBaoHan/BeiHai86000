
//Device Address define
#define   MAX_DEV_LEN   64	//Maximum Length

#define  RX_QUEU    1024 //Receive Buffer size 
#define  TX_QUEU    512  //Send Buffer size 

//Packet Type
#define  PT_READ		0
#define  PT_WRITE		1
#define  PT_READWRITE   2

#if defined(_WIN32_WCE_CEPC) || defined(_ARM_)
typedef  TCHAR *   STR;
#else
typedef  ATOM  STR; 
#endif

union DDEValue
{
	long  lValue;
	float fValue;
	BOOL  bValue;
	STR	  sValue;
};

union PlcValue
{
	BOOL  bitVal;
	BYTE  byteVal;
	short intVal;
	UINT  uintVal;
	WORD  bcdVal;
	long  longVal;
	float floatVal;
	DWORD longbcdVal;
	STR	  strVal;
};

#define BIT_DATATYPE       0x01
#define BYTE_DATATYPE      0x02
#define INT_DATATYPE       0x04
#define UINT_DATATYPE      0x08
#define BCD_DATATYPE       0x10
#define LONG_DATATYPE      0x20
#define LONGBCD_DATATYPE   0x40
#define FLOAT_DATATYPE     0x80
#define STRING_DATATYPE    0x100

#define COM_QUALITY_BAD             0x00
#define COM_QUALITY_UNCERTAIN       0x40
#define COM_QUALITY_GOOD            0xC0
#define COM_QUALITY_CONFIG_ERROR    0x04
#define COM_QUALITY_NOT_CONNECTED   0x08
#define COM_QUALITY_DEVICE_FAILURE  0x0c
#define COM_QUALITY_SENSOR_FAILURE  0x10
#define COM_QUALITY_LAST_KNOWN      0x14
#define COM_QUALITY_COMM_FAILURE    0x18
#define COM_QUALITY_OUT_OF_SERVICE  0x1C
#define COM_QUALITY_LAST_USABLE     0x44
#define COM_QUALITY_SENSOR_CAL      0x50
#define COM_QUALITY_EGU_EXCEEDED    0x54
#define COM_QUALITY_SUB_NORMAL      0x58
#define COM_QUALITY_LOCAL_OVERRIDE  0xD8

typedef struct 
{
	int    nFlag[2];//Features characters 
	TCHAR   sComment[8];//Brief description 
	int    nVersion[2];//Version 
	int    nVarNum;    //variable number
	int    nCommConfig;//Communication config flag
	int    reserved[24];//Reserved
}CFG_FILE_HEAD;


typedef struct
{
	TCHAR     szDevName[32];//Device name 
	TCHAR     szRegister[128];//regsiter name
	short    nDataType; 
	short    nAccessMode;
	short    reserved[2];
}MiniDbItem,MINIDBITEM;

struct ComDevice			//size: 54
{
	short nFlag;			//serial com flag
	WORD  deviceKind;		//device kind, '1' representate serial com device
	TCHAR  deviceName[32];	//Logical name =={COM1£¨COM2...}
	int   baudRate;			//Baud
	short dataBits;			//data bits
	short stopBits;			//stop bits
	short parity;			//parity
	short timeOut;			//timeOut
	short failTimes;		//fail Times
	short retryInterval;	//try connect interval 
	short isRs232;			// RS232 or RS485
};


typedef struct _DevAddr//device address Structure 
{
	UINT	nDevAddr;
	TCHAR	sDevAddr[MAX_DEV_LEN];//device address name
}DEVADDR, * PDEVADDR;


typedef struct Device
{
	TCHAR	szDevName[32];
	TCHAR	szDevAddr[32];
	DEVADDR devAddress;
	
	int		nCommId;
	int		nFailTime;
	BOOL	bBad;
	UINT	nTimerCount;
}DEVICE, * PDEVICE;


typedef struct var
{                 
	WORD wVarID;      			// variable ID 
   	WORD wVarType;      		// variable type
   	TCHAR szVarName[32]; 		// variable name 
	
   	short  nDeviceIndex;		// PLC index        
	short  nUnitNo;				// PLC addressµÿ÷∑
	DEVADDR *pDevAddr;			// pointer to device address structure
   	
	TCHAR   *pszRegName;			//register name
	
	short  nRegType;     		// register type    
	short  nSubType;			// sub-type
	short  nSubType1;
	short  nSubType2;			// 
   	short  nNo;          		// address No.            
	
   	short  nDataType;    		// data type 
	
   	short  nAccessMode;      	// I/O Mode.
	
   	long   nFrequency;   		// Sampling frequency 
   	long  nTimerCount;  		// Counter
    
	CWinThread* pComThread;     // thread 
	DDEValue   maxRaw;			// Maximum  raw value
	DDEValue   minRaw;			// minimum raw value
	short      bConvertion;		// convert  type
	short      isBad;			//bad device
	short      isUnvalid;		//invalid variable	
}PLCVAR, * PPLCVAR;


//packet structure
typedef struct Packet 
{ 
	short  nDeviceIndex;		//-----------//	
	short  nUnitNo;				// device No. 
	DEVADDR * pDevAddr;			// pointer to device structure
	TCHAR    * pszDevName;		// device name 
	TCHAR   *pszRegName;		//regsiter name
	short  nRegType;			// regsiter type
	short  nSubType;			// sub-type
	short  nSubType1;
	short  nSubType2;			// 
	int    nStartNo;			// start offset
	int    nEndNo;				// end offset
    
	short  nPacketType;			// Read, write
	CPtrList varList;			// store 'struct id_no'ªÚ'struct ID_NO2'
}PACKET,* PPACKET;

typedef struct IdNo
{
	WORD wVarId;
	WORD wNo;
	WORD wDataType;
	BYTE Special[4];
	PlcValue  plcValue;
}ID_NO;

typedef struct IdNo2
{
	WORD wVarId;
	WORD wNo;
	WORD wDataType;
	
	PlcValue  plcValue;
	BYTE Special[4];
	
    WORD wQualities;//quality stamps
    FILETIME ftTimeStamps;//time stamps
	
}ID_NO2;


typedef  struct reginfo
{
	TCHAR  sRegName[8];//regsiter name
	
	int   nLowIndex;
	int   nUpperIndex;
	
	WORD  wDataType;//data type format
	int   nData;
}REG_INFO;

typedef  struct deviceinfo //
{
	TCHAR  sDeviceName[32];//Device name
}DEVICE_INFO;

