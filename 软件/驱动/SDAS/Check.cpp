/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Check.cpp
Class Name:  CCheck
Brief:       Check functions class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/

#include "stdafx.h"
#include "Check.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCheck::CCheck()
{

}

CCheck::~CCheck()
{

}

/****************************************************************************
*   Name
		LRC_Check
*	Type
		public
*	Function
		Sum up data in buffer by byte, then inverse the result and add 1 to it
*	Return Value
		one-byte data.
*	Parameters
		buf 
			[in]  the pointer to data buffer
		nLen
			[in]  the length of data 
*****************************************************************************/
BYTE CCheck::LRC_Check(BYTE *buf, int nLen)
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	BYTE checknum=0;
	for (int i=0;i<nLen;i++)
	{
		checknum+=buf[i];
	}
	checknum=~checknum;
	checknum+=1;
	checknum&=0xff;
	return checknum;
}

/****************************************************************************
*   Name
		Sum_Check
*	Type
		public
*	Function
		Sum up data in buffer by byte and get a low 8-bit value of the result
*	Return Value
		one-byte data.
*	Parameters
		buf 
			[in]  the pointer to data buffer
		nLen
			[in]  the length of data 
*****************************************************************************/
BYTE CCheck::Sum_Check(BYTE *buf, int nLen)
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	BYTE checkSum = 0;
	for (int i=0; i<nLen; i++)
	{
		checkSum += (BYTE)buf[i];
	}
	return checkSum;
}

/****************************************************************************
*   Name
		Xor_Check
*	Type
		public
*	Function
		XOR(exclusive or) data in buffer by byte and get a low 8-bit value of the result
*	Return Value
		one-byte data.
*	Parameters
		buf 
			[in]  the pointer to data buffer
		nLen
			[in]  the length of data 
*****************************************************************************/
BYTE CCheck::Xor_Check(BYTE *buf, int nLen)
{
	ASSERT(buf != NULL);
	
	if (buf == NULL)
	{
		return 0;
	}

	BYTE checkSum = (BYTE)*buf;
	for(int i=1; i<nLen; i++)
	{
		checkSum ^= (BYTE)buf[i]; 
	}
	return checkSum;
}

//2^bitWidth - 1. e.g. widthMask(16) = 0xFFFF.
inline long widthMask(int bitWidth)
{
	return (((1L << (bitWidth - 1)) - 1L) << 1) | 1L;
}
//inverse the low  itnumbit, For example,  reflect(0x3e23,3) = 0x3e26
inline long reflect(long value, int numBits) {
	long temp = value;

	for (int i = 0; i < numBits; i++) {
		long bitMask = 1L << ((numBits - 1) - i);

		if ((temp & 1L) != 0) {
			value |= bitMask;
		}
		else {
			value &= ~bitMask;
		}		
		temp >>= 1;
	}

	return value;
}

/****************************************************************************
*   Name
		CRC_Check
*	Type
		public
*	Function
		Cyclic Redundancy Check
*	Return Value
		a long integer data
*	Parameters
		buf 
			[in]   the pointer to data buffer
		iLen 
			[in]   buffer size
		bitWidth 
			[in]   the bit width, =8,16,32,64
		polynomial 
			[in]   CRCPolynomial
		reflectPolynomial 
			[in]   inverse or not
		initialCrcValueValue 
			[in]   crc Initial value
		reflectInput 
			[in]   reflect Input byte
		reflectOupt 
			[in]   reflect Ouput byte
		xorOutputValue 
			[in]   the last CRC mask
*	Remarks:
		calculate Method:   
			CRC4=X4+X1+1           //0x03,0011
			CRC8=X8+X5+X4+1        //0x31,00110001
			CRC12=X12+X11+X3+X2+1  //0x080D,00001000 00001101
			CRC-CCITT=X16+X12+X5+1 //0x1021,00010000 00100001
			CRC16=X16+X15+X2+1     //0x8005,10000000 00000101    
			CRC32=X32+X26+X23+X22+X16+X12+X11+X10+X8+X7+X5+X4+X2+X1+1 //0x04C11DB7,00000100 11000001 00011101 10110111
*****************************************************************************/
long CCheck::CRC_Check(BYTE *buf,
			   int iLen,
			   int bitWidth, 
			   long polynomial,
			   bool reflectPolynomial,
			   long initiallCrcValueValue/* = 0*/, 
			   bool reflectInput/* = 0*/,
			   bool reflectOutput/* = 0*/,
			   long xorOutputValue/* = 0*/)
{
	//initial
	long lCrcValue = initiallCrcValueValue;

	if (!reflectPolynomial)
	{
		long topBit = 1L << (bitWidth - 1);
		for (int ix = 0; ix < iLen; ix++)
		{
			//CRC main Process 
			if (reflectInput) {
				buf[ix] = static_cast<BYTE>(reflect(buf[ix], 8));
			}
			
			lCrcValue ^= (buf[ix] << (bitWidth - 8));

			for (int i = 0; i < 8; i++) {
				if ((lCrcValue & topBit) != 0) {
					lCrcValue = (lCrcValue << 1) ^ polynomial;
				}
				else {
					lCrcValue <<= 1;
				}
				//
				lCrcValue &= widthMask(bitWidth);
			}		
		}
	}
	else
	{
		long topBit = 1L;
		polynomial = reflect(polynomial, bitWidth);
		for (int ix = 0; ix < iLen; ix++)
		{
			//CRC main Process 			
			if (reflectInput) {
				buf[ix] = static_cast<BYTE>(reflect(buf[ix], 8));
			}

			lCrcValue ^= buf[ix] ;

			for (int i = 0; i < 8; i++) {
				if ((lCrcValue & topBit) != 0) {
					lCrcValue = (lCrcValue >> 1) ^ polynomial;
				}
				else {
					lCrcValue >>= 1;
				}
				lCrcValue &= widthMask(bitWidth);
			}		
		}
	}
	
	lCrcValue = xorOutputValue ^ lCrcValue;

	if (reflectOutput) {
		lCrcValue = xorOutputValue ^ reflect(lCrcValue, bitWidth);
	}
	return lCrcValue;
}

/****************************************************************************
*   Name
		CRC16_CHECK_REVERSE
*	Type
		public
*	Function
		16-bits Cyclic Redundancy Check (reverse every bit)
*	Return Value
		a long integer data.
*	Parameters
		buf 
			[in]  the pointer to data buffer
		nLen
			[in]  the length of data 
*****************************************************************************/
WORD CCheck::CRC16_CHECK_REVERSE(BYTE *buf, int nLen)
{
	return (WORD)CRC_Check(buf, nLen, 16, CRC16, true, 0xFFFF);
	
}

/****************************************************************************
*   Name
		CRC16_CHECK
*	Type
		public
*	Function
		16-bits Cyclic Redundancy Check 
*	Return Value
		a long integer data.
*	Parameters
		buf 
			[in]  the pointer to data buffer
		nLen
			[in]  the length of data 
*****************************************************************************/
WORD CCheck::CRC16_CHECK(BYTE *buf, int nLen)
{
	return (WORD)CRC_Check(buf, nLen, 16, CRC16, 0);
	
}

/****************************************************************************
*   Name
		MODBUS_CRC_CHECK
*	Type
		public
*	Function
		Mobus protocol Cyclic Redundancy Check 
*	Return Value
		a long integer data.
*	Parameters
		buf 
			[in]  the pointer to data buffer
		nLen
			[in]  the length of data 
*****************************************************************************/
WORD CCheck::MODBUS_CRC_CHECK(BYTE *buf, int nLen)
{
	return CRC16_CHECK_REVERSE(buf, nLen);
}