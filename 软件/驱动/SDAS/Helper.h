/************************************************
Copyright(c) 2006, AsiaControl Company
All rights reserved.
File Name:   Helper.h
Class Name:  CHelper
Brief:       Help functions class
History: 
Date	         Author  	       Remarks
Aug. 2006	     Develop Dep.1     
************************************************/
#pragma once

class CHelper
{
public:
	CHelper(void);
	~CHelper(void);
	
public:
	
	static ULONGLONG KvGetTickCount();
	static WORD GetWord(BYTE* buf);
	static long GetLong(BYTE* buf);
	static WORD WordToBcd(WORD value);
	static long LongToBcd(long value);
	static WORD BcdToWord(WORD value);
	static long BcdToLong(long value);
	static WORD SwapWord(WORD value);
	static long SwapLong(WORD *value);
	static float ByteBufToFloat(BYTE *buf,int mode);
	static void FloatToByteBuf(BYTE *buf,float data,int mode);
	static void ByteToASC( BYTE* buf, BYTE data );
	static BYTE ASCToByte( BYTE* buf );
	static void WordToASC( BYTE* buf, WORD data );
	static WORD ASCToWord( BYTE* buf );
	static void LongToASC(BYTE *buf ,long data);
	static LONG ASCToLong(BYTE *buf);
	static void BCDToASC( BYTE* buf, WORD data );
	static WORD ASCToBCD( BYTE* buf );

	static void KvCoFileTimeNow( FILETIME FAR* lpFileTime);
	static UINT KvGlobalGetAtomName(STR  pAtom, LPTSTR lpBuffer, int nSize);
	static STR  KvGlobalDeleteAtom(STR  pAtom);
	static STR  KvGlobalAddAtom(LPCTSTR lpString);
	
};
