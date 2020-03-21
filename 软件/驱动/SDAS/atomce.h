//AtomCE.h
#ifndef __ATOMCE_H__
#define __ATOMCE_H__
// X86 CPU and ARM use different library. 
#if defined(_X86_) 
#pragma comment(lib, "C:\\Program Files\\KingViewDriver\\lib\\FORX86LIB\\AtomCE.lib")
#elif defined(_ARM_) 
#pragma comment(lib, "C:\\Program Files\\KingViewDriver\\lib\\FORARMV4LIB\\AtomCE.lib")
#endif


/****************************************************************************
*   Name
		MyGlobalAddAtom
*	Type
		public
*	Function
		adds a character string to the global atom table and returns a unique value
*	Return Value
		a atom
*	Parameters
		lpString 
		[in] Pointer to the null-terminated string that you want to add to the 
		global atom table
*****************************************************************************/
__declspec(dllimport) TCHAR* MyGlobalAddAtom(LPCTSTR lpString);

/****************************************************************************
*   Name
		MyGlobalDeleteAtom
*	Type
		public
*	Function
		decrements the reference count of a global string atom
*	Return Value
		a atom. 
*	Parameters
		pAtom
			[in] Atom that identifies the character string that you want to
			delete or for which you want to decrement the reference count. 

*****************************************************************************/
__declspec(dllimport) TCHAR* MyGlobalDeleteAtom(TCHAR* pAtom);

/****************************************************************************
*   Name
		MyGlobalGetAtomName
*	Type
		public
*	Function
		retrieves a copy of the character string associated with the specified global atom
*	Return Value
		Return Value is the length of the string copied to the buffer, in TCHARs  
*	Parameters
		pAtom
			[in] Identifies the global atom associated with the character string to be 
			retrieved.
		lpBuffer
			[out] Pointer to the buffer for the character string
		nSize
			[in] Specifies the size, in TCHARs, of the buffer
*****************************************************************************/

__declspec(dllimport) UINT MyGlobalGetAtomName(TCHAR* pAtom, LPTSTR lpBuffer, int nSize);

#endif//__ATOMCE_H__
