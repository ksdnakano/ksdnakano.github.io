#include "StdAfx.h"
#include "ExcelEvents.h"

/******************************************************************************
*   Invoke -- Takes a dispid and uses it to call another of this class's 
*   methods. Returns S_OK if the call was successful.
******************************************************************************/ 
STDMETHODIMP CExcelEvents::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                  WORD wFlags, DISPPARAMS* pDispParams,
                                  VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                  UINT* puArgErr)
{
   switch(dispIdMember){
/*   case 0x00622:
      if(pDispParams->cArgs !=2)
         return E_INVALIDARG;
      else
      {
         if(pDispParams->rgvarg[1].vt & VT_BYREF)
         {
            OnBeforeWorkbookClose( // Call the function.
               *(pDispParams->rgvarg[1].ppdispVal),
               pDispParams->rgvarg[0].pboolVal);
         }
         else
         {
            OnBeforeWorkbookClose(  // Call the function.
               (pDispParams->rgvarg[1].pdispVal),
               pDispParams->rgvarg[0].pboolVal);
         }
      }
		break;*/
   case	DISPID_EXCELAPP_NEWWORKBOOK			:
   case	DISPID_EXCELAPP_WORKBOOKACTIVATE	:
   case	DISPID_EXCELAPP_WORKBOOKOPEN		:
   case	DISPID_EXCELAPP_SHEETACTIVATE		:

	   break;
   case	DISPID_EXCELAPP_SHEETSELECTIONCHANGE	:
		TRACE( _T("Event %08x\r\n"),dispIdMember );
		break;
   case DISPID_EXCELAPP_SHEETCHANGE:
      {
         if(pDispParams->rgvarg[1].vt & VT_BYREF)
         {
            OnSheetChange( // Call the function.
               *(pDispParams->rgvarg[1].ppdispVal),
               *(pDispParams->rgvarg[0].ppdispVal));
         }
         else
         {
            OnSheetChange(  // Call the function.
               pDispParams->rgvarg[1].pdispVal,
               pDispParams->rgvarg[0].pdispVal);
         }
      }
      break;
   }
   return S_OK;
}

/******************************************************************************
*  HandleSheetChange -- This method processes the SheetChange event for the 
*  application attached to this event handler.
******************************************************************************/ 
STDMETHODIMP CExcelEvents::OnSheetChange( IDispatch* xlSheet, 
                                                  IDispatch* xlRange)
{
   HRESULT hr = S_OK;
   return hr;
}
