	typedef enum ICC_CONSTANTS {
		ICC_LISTVIEW_CLASSES = 0x00000001, // listview, header
		ICC_TREEVIEW_CLASSES = 0x00000002, // treeview, tooltips
		ICC_BAR_CLASSES = 0x00000004, // toolbar, statusbar, trackbar, tooltips
		ICC_TAB_CLASSES = 0x00000008, // tab, tooltips
		ICC_UPDOWN_CLASS = 0x00000010, // updown
		ICC_PROGRESS_CLASS = 0x00000020, // progress
		ICC_HOTKEY_CLASS = 0x00000040, // hotkey
		ICC_ANIMATE_CLASS = 0x00000080, // animate
		ICC_WIN95_CLASSES = 0x000000FF,
		ICC_DATE_CLASSES = 0x00000100, // month picker, date picker, time picker, updown
		ICC_USEREX_CLASSES = 0x00000200, // comboex
		ICC_COOL_CLASSES = 0x00000400, // rebar (coolbar) control
		ICC_INTERNET_CLASSES = 0x00000800,
		ICC_PAGESCROLLER_CLASS = 0x00001000,   // page scroller
		ICC_NATIVEFNTCTL_CLASS = 0x00002000,   // native font control
		ICC_STANDARD_CLASSES = 0x00004000,
		ICC_LINK_CLASS = 0x00008000
	} ICC_CONSTANTS;

	typedef struct tagINITCOMMONCONTROLSEX {
		long dwSize;
		ICC_CONSTANTS dwICC;
	} tagINITCOMMONCONTROLSEX;

	[dllname("COMCTL32.DLL")]
	module comctl32 {

		[entry("InitCommonControlsEx")]
		long InitCommonControlsEx(
			[in] tagINITCOMMONCONTROLSEX* tICC
		);

		[entry("SetWindowSubclass")]
		long SetWindowSubclass(
			[in] long hwnd,
			[in] long pfnSubclass,
			[in] long uIdSubclass,
			[in] void* dwRefData
		);

		[entry("RemoveWindowSubclass")]
		long RemoveWindowSubclass(
			[in] long hwnd,
			[in] long pfnSubclass,
			[in] long uIdSubclass
		);

		[entry("DefSubclassProc")]
		long DefSubclassProc(
			[in] long hwnd,
			[in] long uMsg,
			[in] long wParam,
			[in] void* lParam
		);

	}

	#include "edit_ctl.inc"