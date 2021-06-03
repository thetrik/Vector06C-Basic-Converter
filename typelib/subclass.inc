
[uuid(DEEB842B-4D58-44BE-AC71-E1D2EECF1591)]
interface ISubclass : IUnknown {

	[propget]
	HRESULT hWnd(
		[out, retval] long* pResult);

	HRESULT OnWindowProc(
		[in] long hWnd,
		[in] long lMsg,
		[in] long wParam,
		[in] long lParam,
		[in, out] VARIANT_BOOL *bDefCall,
		[out, retval] long* pResult);

}