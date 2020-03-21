/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Helper.cpp
Class Name:  CHelper
Brief:       Help functions class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#include "StdAfx.h"
#include "helper.h"
#include "math.h"

CHelper::CHelper(void)
{
}

CHelper::~CHelper(void)
{
}


/*=============================================================
	Data conversion functions
===============================================================*/

/****************************************************************************
*   Name
		GetWord
*	Type
		public
*	Function
		Get an integer from a buffer
*	Return Value
		an integer
*	Parameters
		buf 
			[in]   the pointer to data buffer
*****************************************************************************/
WORD CHelper::GetWord(BYTE* buf)
{
	ASSERT(buf != NULL);
	if (buf == NULL)
	{
		return 0;
	}

	WORD value = buf[0]*256+buf[1];
	return value;
}

/****************************************************************************
*   Name
		GetLong
*	Type
		public
*	Function
		Get a long integer from a buffer
*	Return Value
		a long integer 
*	Parameters
		buf 
			[in]   the pointer to data buffer
*****************************************************************************/
long CHelper::GetLong(BYTE* buf)
{
	ASSERT(buf != NULL);
	if (buf == NULL)
	{
		return 0;
	}

	long value=0;
	WORD temp1 = GetWord(buf);
	WORD temp2 = GetWord(buf+2);
	value = temp1*65536+temp2;
	return value;
}

/****************************************************************************
*   Name
		WordToBcd
*	Type
		public
*	Function
		Convert an integer into BCD format
*	Return Value
		a BCD-format number 
*	Parameters
		value 
			[in]   an integer to convert
*****************************************************************************/
WORD CHelper::WordToBcd(WORD value)
{
	WORD data =0;
	for(int i=0;i<4;i++)
	{
		data = data*10+(WORD)(value/pow(16,3-i));
		WORD d = (WORD)pow(16,3-i);
		value = (WORD)(value%d);
	}
	return data;

}

/****************************************************************************
*   Name
		LongToBcd
*	Type
		public
*	Function
		Contert a long integer into BCD format
*	Return Value
		a BCD-format number 
*	Parameters
		value 
			[in]  a long integer to convert 
*****************************************************************************/
long CHelper::LongToBcd(long value)
{
	unsigned long data =0;
	for(int i=0;i<8;i++)
	{
		data = data*10+(long)(value/pow(16,7-i));
		unsigned long d = (long)pow(16,7-i);
		value = (unsigned long)(value%d);
	}
	return data;

}

/****************************************************************************
*   Name
		BcdToWord
*	Type
		public
*	Function
		Convert BCD-format number into an integer
*	Return Value
		an integer
*	Parameters
		value 
			[in]    a BCD-format number to convert
*****************************************************************************/
WORD CHelper::BcdToWord(WORD value)
{
	WORD data = 0;
	WORD kk =0;
	for(int i=0;i<4;i++)
	{
		data = data*16+(WORD)(value/pow(10,3-i));
		kk = (WORD)pow(10,3-i);
        value = (WORD)value%kk;
	}
	return data;
}

/****************************************************************************
*   Name
		BcdToLong
*	Type
		public
*	Function
		Convert a BCD-format number into a long integer
*	Return Value
		a long integer
*	Parameters
		value 
			[in]   a BCD-format number to convert
*****************************************************************************/
long CHelper::BcdToLong(long value)
{
	long data = 0;
	long kk =0;
	for(int i=0;i<8;i++)
	{
		data = data*16+(WORD)(value/pow(10,7-i));
		kk = (WORD)pow(10,7-i);
        value = (WORD)value%kk;
	}
	return data;
}

/****************************************************************************
*   Name
		SwapWord
*	Type
		public
*	Function
		Swap the high and low bytes of an integer
*	Return Value
		an integer 
*	Parameters
		value 
			[in]  an integer to swap
*****************************************************************************/
WORD CHelper::SwapWord(WORD value)
{
	WORD temp=0;
	temp=value;
	WORD wData;
	wData = (value>>8)+(temp<<8);
	return wData;
}

/****************************************************************************
*   Name
		SwapLong
*	Type
		public
*	Function
		Swap the high and low bytes of a long integer
*	Return Value
		a long integer  
*	Parameters
		value 
			[in]  a long integer to swap
*****************************************************************************/
long CHelper::SwapLong(WORD *value)
{
	ASSERT(value != NULL);
	if (value == NULL)
	{
		return 0;
	}

	WORD temp = SwapWord(value[0]);
	WORD wData = SwapWord(value[1]);
	long lvalue = wData*65536+temp;
	return lvalue;
}

/****************************************************************************
*   Name
		ByteBufToFloat
*	Type
		public
*	Function
		Convert a 4-byte data into FLOAT type according mode
*	Return Value
		a floating-point number 
*	Parameters
		value 
			[in]  the pointer to data buffer
		mode 
			[in]  data conversion mode. The following is acceptable. 
									¡¡	    0 --3210
											1 --2301
											2 --0123
											3 --1032
*****************************************************************************/
float CHelper::ByteBufToFloat(BYTE *buf,int mode)
{	
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	BYTE *pBuf=buf;
	BYTE buff[4];
	switch(mode)
	{
	default:
	case 0: 
		buff[0]=pBuf[3];
		buff[1]=pBuf[2];
		buff[2]=pBuf[1];
		buff[3]=pBuf[0];
		break;
	case 1: 
		buff[0]=pBuf[2];
		buff[1]=pBuf[3];
		buff[2]=pBuf[0];
		buff[3]=pBuf[1];
		break;
	case 2: 
		buff[0]=pBuf[0];
		buff[1]=pBuf[1];
		buff[2]=pBuf[2];
		buff[3]=pBuf[3];
		break;
	case 3: 
		buff[0]=pBuf[1];
		buff[1]=pBuf[0];
		buff[2]=pBuf[3];
		buff[3]=pBuf[2];
		break;
	}
	float Data;
	memcpy(&Data,buff,4);
	return  Data;
}

/****************************************************************************
*   Name
		FloatToByteBuf
*	Type
		public
*	Function
		Convert a FLOAT-type number into 4-byte format according mode
*	Return Value
		a floating-point number 
*	Parameters
		buf 
			[out]   the pointer to data buffer
		data 
			[in]   a FLOAT-type number to convert 
		mode 
			[in]  data conversion mode. The following is acceptable. 
									¡¡	    0 --3210
											1 --2301
											2 --0123
											3 --1032
*****************************************************************************/
void CHelper::FloatToByteBuf(BYTE *buf,float data,int mode)
{
	BYTE temp[10];
	memcpy(temp,&data,4);
	switch(mode)
	{
	case 0:
		buf[0] = temp[3];
		buf[1] = temp[2];
		buf[2] = temp[1];
		buf[3] = temp[0];
		break;
	case 1:
		buf[0] = temp[2];
		buf[1] = temp[3];
		buf[2] = temp[0];
		buf[3] = temp[1];
		break;
	case 2:
		buf[0] = temp[0];
		buf[1] = temp[1];
		buf[2] = temp[2];
		buf[3] = temp[3];
		break;
	case 3:
		buf[0] = temp[1];
		buf[1] = temp[0];
		buf[2] = temp[3];
		buf[3] = temp[2];
		break;
	default:
		break;
	}
}

/****************************************************************************
*   Name
		ByteToASC
*	Type
		public
*	Function
		Convert a one-byte data into two-ASCII format and store it in buf.
*	Return Value
		a long integer  
*	Parameters
		buf 
			[out]   the pointer to data buffer for saving the result
		data 
			[out]    a one-byte data to convert 
*****************************************************************************/
void CHelper::ByteToASC( BYTE* buf, BYTE data )
{	
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return ;
	}

    BYTE b = data>>4;
    if ( b < 0x0A )
        buf[0] = b + '0';
    else  // A~F
        buf[0] = b - 0x0A + 'A';
    
    b = data&0x0F;
    if ( b < 0x0A )
        buf[1] = b + '0';
    else  // A~F
        buf[1] = b - 0x0A + 'A';
}

/****************************************************************************
*   Name
		ByteToASC
*	Type
		public
*	Function
		Convert a two-ASCII data into a one-byte data.
*	Return Value
		a one-byte data 
*	Parameters
		buf 
			[in]   the pointer to data buffer
*****************************************************************************/
BYTE CHelper::ASCToByte( BYTE* buf )
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	BYTE b = 0;
    if ( buf[0] <= '9' && buf[0] >= '0' )
        b = (buf[0] - '0') << 4;
    else  // A~F
        b = (buf[0] - 'A' + 0x0A) << 4;
    if ( buf[1] <= '9' && buf[1] >= '0' )
        b += (buf[1] - '0');
    else  // A~F
        b += (buf[1] - 'A' + 0x0A);
    return b;
}

/****************************************************************************
*   Name
		ByteToASC
*	Type
		public
*	Function
		Convert an one-word data into 4-ASCII format
*	Return Value
		null  
*	Parameters
		buf 
			[out]   the pointer to data buffer for saving the result
		data 
			[in]    a one-byte data to convert 
*****************************************************************************/
void CHelper::WordToASC( BYTE* buf, WORD data )
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return ;
	}


	ByteToASC( buf, HIBYTE(data) );
    ByteToASC( buf + 2, LOBYTE(data) );
}

/****************************************************************************
*   Name
		ASCToWord
*	Type
		public
*	Function
		Convert a 4-ASCII data into an one-word format
*	Return Value
		an one-word number
*	Parameters
		buf 
			[in]   the pointer to data buffer
*****************************************************************************/
WORD CHelper::ASCToWord( BYTE* buf )
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

    WORD data = ((WORD)ASCToByte(buf)) | (WORD)ASCToByte(buf+2)<<8;
    return data;
}

/****************************************************************************
*   Name
		LongToASC
*	Type
		public
*	Function
		Convert a double-word data into 8-ASCII format
*	Return Value
		null
*	Parameters
		buf 
			[out]   the pointer to data buffer for saving the result
		data 
			[in]    a long integer to convert
*****************************************************************************/
void CHelper::LongToASC(BYTE *buf ,long data)
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return ;
	}


	WordToASC( buf, HIWORD(data) );
    WordToASC( buf + 4, LOWORD(data) );
}

/****************************************************************************
*   Name
		ASCToLong
*	Type
		public
*	Function
		Convert a 8-ASCII data into double-word format
*	Return Value
		an long integer
*	Parameters
		buf 
			[in]   the pointer to data buffer
*****************************************************************************/
LONG CHelper::ASCToLong(BYTE *buf)
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	long data = ((long)ASCToWord(buf)) | (long)ASCToWord(buf+4)<<16;
    return data;
}

/****************************************************************************
*   Name
		BCDToASC
*	Type
		public
*	Function
		Convert an one-word BCD data into 4-ASCII format
*	Return Value
		null 
*	Parameters
		buf 
			[out]   the pointer to data buffer for saving the result
		data 
			[in]    a BCD-format number to convert 
*****************************************************************************/
void CHelper::BCDToASC( BYTE* buf, WORD data )
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return ;
	}

	WORD wData = data;
    for ( int i = 0; i < 4; i ++ ) 
	{
        buf[3-i] = (BYTE)(wData % 10) + '0';
        wData /= 10;
    }
}

/****************************************************************************
*   Name
		ASCToBCD
*	Type
		public
*	Function
		Convert a 4-ASCII data into one-word BCD format
*	Return Value
		an integer 
*	Parameters
		buf 
			[in]   the pointer to data buffer
*****************************************************************************/
WORD CHelper::ASCToBCD( BYTE* buf )
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	WORD data = 0;
    for ( int i = 0; i < 4; i ++ )
	{
        data *= 10;
        data += buf[i] - '0';
    }
    return data;
}

/****************************************************************************
*   Name
		KvGetTickCount
*	Type
		public
*	Function
		Get the current time (ms)
*	Return Value
		an unsigned longlong  
*	Parameters
		null
*****************************************************************************/
ULONGLONG CHelper::KvGetTickCount()
{
	FILETIME ft;
	SYSTEMTIME st;
	GetSystemTime(&st); // gets current time
	SystemTimeToFileTime(&st, &ft); // converts to file time format
	ULONGLONG uTimeCount =((((ULONGLONG) ft.dwHighDateTime) << 32) + ft.dwLowDateTime)/10000;
	
	return uTimeCount;
}

/****************************************************************************
*   Name
		KvCoFileTimeNow
*	Type
		public
*	Function
		Setting the current time
*	Return Value
		null  
*	Parameters
		lpFileTime
			[out] a FILETIME pointer
*****************************************************************************/
void CHelper::KvCoFileTimeNow(FILETIME *lpFileTime) //
{
#if defined(_WIN32_WCE_CEPC) || defined(_ARM_)
	CTime time(CTime::GetCurrentTime());
	SYSTEMTIME timeDest;
	time.GetAsSystemTime(timeDest);
	SystemTimeToFileTime(&timeDest, lpFileTime);
#else
	CoFileTimeNow(lpFileTime);	
#endif	 
}

/****************************************************************************
*   Name
		KvGlobalAddAtom
*	Type
		public
*	Function
		adds a character string to the global atom table and returns a unique value
*	Return Value
		a STR variable.   
*	Parameters
		lpString 
			[in] Pointer to the null-terminated string that you want to add to the 
			global atom table,
*****************************************************************************/
STR  CHelper::KvGlobalAddAtom(LPCTSTR lpString)
{
#if defined(_WIN32_WCE_CEPC) || defined(_ARM_)
	return MyGlobalAddAtom(lpString);	
#else
	return GlobalAddAtom(lpString);
#endif
}

/****************************************************************************
*   Name
		KvGlobalGetAtomName
*	Type
		public
*	Function
		retrieves a copy of the character string associated with the specified global 
		atom
*	Return Value
		 return value is the length of the string copied to the buffer, in TCHARs 
*	Parameters
		pAtom
			[in] Identifies the global atom associated with the character string to 
			be retrieved.
		lpBuffer
			[out] Pointer to the buffer for the character string
		nSize
			[in] Specifies the size, in TCHARs, of the buffer
*****************************************************************************/
UINT CHelper::KvGlobalGetAtomName(STR  pAtom, LPTSTR lpBuffer, int nSize)
{
#if defined(_WIN32_WCE_CEPC) || defined(_ARM_)
	return MyGlobalGetAtomName(pAtom , lpBuffer,nSize);
#else
	return GlobalGetAtomName(pAtom , lpBuffer,nSize);
#endif
}

/****************************************************************************
*   Name
		KvGlobalDeleteAtom
*	Type
		public
*	Function
		decrements the reference count of a global string atom
*	Return Value
		a STR variable. 
*	Parameters
		pAtom
			[in] Atom that identifies the character string that you want to
			delete or for which you want to decrement the reference count. 
*****************************************************************************/
STR  CHelper::KvGlobalDeleteAtom(STR  pAtom)
{
#if defined(_WIN32_WCE_CEPC) || defined(_ARM_)
	return MyGlobalDeleteAtom( pAtom);
#else
	return GlobalDeleteAtom( pAtom);
#endif
}

