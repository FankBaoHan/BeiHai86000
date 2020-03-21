/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Check.h
Class Name:  CCheck
Brief:       Check functions class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#pragma once

#define CRC4 0x03
#define CRC8 0x31
#define CRC12 0x080D
#define CRC16 0x8005
#define CRC_ITU 0x1021
#define CRC_CCITT CRC_ITU
#define CRC32 0x04C11DB7	 

class CCheck  
{
public:
	static BYTE Xor_Check(BYTE *buf,int nLen);
	static BYTE Sum_Check(BYTE *buf,int nLen);
	static BYTE LRC_Check( BYTE * buf,int nLen);
	static long CRC_Check(BYTE *buf,
			   int iLen,
			   int bitWidth, 
			   long polynomial,
			   bool reflectPolynomial,
			   long initiallCrcValueValue = 0, 
			   bool reflectInput = 0,
			   bool reflectOutput = 0,
			   long xorOutputValue = 0);
	static WORD MODBUS_CRC_CHECK(BYTE *buf, int nLen);
	static WORD CRC16_CHECK(BYTE *buf, int nLen);
	static WORD CRC16_CHECK_REVERSE(BYTE *buf, int nLen);
	CCheck();
	virtual ~CCheck();

};
