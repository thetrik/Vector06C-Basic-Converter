	typedef struct POINT {
		long x;
		long y;
	} POINT;

	typedef struct RECT {
		long Left;
		long Top;
		long Right;
		long Bottom;
	} RECT;

	typedef enum CLIPBOARD_FORMAT_CONSTS {
		CF_TEXT = 1,
		CF_BITMAP = 2,
		CF_METAFILEPICT = 3,
		CF_SYLK = 4,
		CF_DIF = 5,
		CF_TIFF = 6,
		CF_OEMTEXT = 7,
		CF_DIB = 8,
		CF_PALETTE = 9,
		CF_PENDATA = 10,
		CF_RIFF = 11,
		CF_WAVE = 12,
		CF_UNICODETEXT = 13,
		CF_ENHMETAFILE = 14,
		CF_HDROP = 15,
		CF_LOCALE = 16,
		CF_DIBV5 = 17,
		CF_MAX = 18,
		CF_OWNERDISPLAY = 0x0080,
		CF_DSPTEXT = 0x0081,
		CF_DSPBITMAP = 0x0082,
		CF_DSPMETAFILEPICT = 0x0083,
		CF_DSPENHMETAFILE = 0x008E,
		CF_PRIVATEFIRST = 0x0200,
		CF_PRIVATELAST = 0x02FF,
		CF_GDIOBJFIRST = 0x0300,
		CF_GDIOBJLAST = 0x03FF
	} CLIPBOARD_FORMAT_CONSTS;

	typedef enum GWL_CONSTS {
		GWL_WNDPROC=(-4),
		GWL_HINSTANCE=(-6),
		GWL_HWNDPARENT=(-8),
		GWL_STYLE=(-16),
		GWL_EXSTYLE=(-20),
		GWL_USERDATA=(-21),
		GWL_ID=(-12),
		DWL_MSGRESULT=0,
		DWL_DLGPROC=4,
		DWL_USER=8
	} GWL_CONSTS;

	typedef enum DRAWTEXT_FORMATS {
		DT_TOP = 0x00000000,
		DT_LEFT = 0x00000000,
		DT_CENTER = 0x00000001,
		DT_RIGHT = 0x00000002,
		DT_VCENTER = 0x00000004,
		DT_BOTTOM = 0x00000008,
		DT_WORDBREAK = 0x00000010,
		DT_SINGLELINE = 0x00000020,
		DT_EXPANDTABS = 0x00000040,
		DT_TABSTOP = 0x00000080,
		DT_NOCLIP = 0x00000100,
		DT_EXTERNALLEADING = 0x00000200,
		DT_CALCRECT = 0x00000400,
		DT_NOPREFIX = 0x00000800,
		DT_INTERNAL = 0x00001000,
		DT_EDITCONTROL = 0x00002000,
		DT_PATH_ELLIPSIS = 0x00004000,
		DT_END_ELLIPSIS = 0x00008000,
		DT_MODIFYSTRING = 0x00010000,
		DT_RTLREADING = 0x00020000,
		DT_WORD_ELLIPSIS = 0x00040000,
		DT_NOFULLWIDTHCHARBREAK = 0x00080000,
		DT_HIDEPREFIX = 0x00100000,
		DT_PREFIXONLY = 0x00200000
	} DRAWTEXT_FORMATS;

	typedef enum WM_CONSTANTS {
		WM_NULL                        = 0x0000,
		WM_CREATE                      = 0x0001,
		WM_DESTROY                     = 0x0002,
		WM_MOVE                        = 0x0003,
		WM_SIZE                        = 0x0005,
		WM_ACTIVATE                    = 0x0006,
		WM_SETFOCUS                    = 0x0007,
		WM_KILLFOCUS                   = 0x0008,
		WM_ENABLE                      = 0x000A,
		WM_SETREDRAW                   = 0x000B,
		WM_SETTEXT                     = 0x000C,
		WM_GETTEXT                     = 0x000D,
		WM_GETTEXTLENGTH               = 0x000E,
		WM_PAINT                       = 0x000F,
		WM_CLOSE                       = 0x0010,
		WM_QUERYENDSESSION             = 0x0011,
		WM_QUERYOPEN                   = 0x0013,
		WM_ENDSESSION                  = 0x0016,
		WM_QUIT                        = 0x0012,
		WM_ERASEBKGND                  = 0x0014,
		WM_SYSCOLORCHANGE              = 0x0015,
		WM_SHOWWINDOW                  = 0x0018,
		WM_WININICHANGE                = 0x001A,
		WM_SETTINGCHANGE               = 0x001A,
		WM_DEVMODECHANGE               = 0x001B,
		WM_ACTIVATEAPP                 = 0x001C,
		WM_FONTCHANGE                  = 0x001D,
		WM_TIMECHANGE                  = 0x001E,
		WM_CANCELMODE                  = 0x001F,
		WM_SETCURSOR                   = 0x0020,
		WM_MOUSEACTIVATE               = 0x0021,
		WM_CHILDACTIVATE               = 0x0022,
		WM_QUEUESYNC                   = 0x0023,
		WM_GETMINMAXINFO               = 0x0024,
		WM_PAINTICON                   = 0x0026,
		WM_ICONERASEBKGND              = 0x0027,
		WM_NEXTDLGCTL                  = 0x0028,
		WM_SPOOLERSTATUS               = 0x002A,
		WM_DRAWITEM                    = 0x002B,
		WM_MEASUREITEM                 = 0x002C,
		WM_DELETEITEM                  = 0x002D,
		WM_VKEYTOITEM                  = 0x002E,
		WM_CHARTOITEM                  = 0x002F,
		WM_SETFONT                     = 0x0030,
		WM_GETFONT                     = 0x0031,
		WM_SETHOTKEY                   = 0x0032,
		WM_GETHOTKEY                   = 0x0033,
		WM_QUERYDRAGICON               = 0x0037,
		WM_COMPAREITEM                 = 0x0039,
		WM_GETOBJECT                   = 0x003D,
		WM_COMPACTING                  = 0x0041,
		WM_COMMNOTIFY                  = 0x0044,
		WM_WINDOWPOSCHANGING           = 0x0046,
		WM_WINDOWPOSCHANGED            = 0x0047,
		WM_POWER                       = 0x0048,
		WM_COPYDATA                    = 0x004A,
		WM_CANCELJOURNAL               = 0x004B,
		WM_NOTIFY                      = 0x004E,
		WM_INPUTLANGCHANGEREQUEST      = 0x0050,
		WM_INPUTLANGCHANGE             = 0x0051,
		WM_TCARD                       = 0x0052,
		WM_HELP                        = 0x0053,
		WM_USERCHANGED                 = 0x0054,
		WM_NOTIFYFORMAT                = 0x0055,
		WM_CONTEXTMENU                 = 0x007B,
		WM_STYLECHANGING               = 0x007C,
		WM_STYLECHANGED                = 0x007D,
		WM_DISPLAYCHANGE               = 0x007E,
		WM_GETICON                     = 0x007F,
		WM_SETICON                     = 0x0080,
		WM_NCCREATE                    = 0x0081,
		WM_NCDESTROY                   = 0x0082,
		WM_NCCALCSIZE                  = 0x0083,
		WM_NCHITTEST                   = 0x0084,
		WM_NCPAINT                     = 0x0085,
		WM_NCACTIVATE                  = 0x0086,
		WM_GETDLGCODE                  = 0x0087,
		WM_SYNCPAINT                   = 0x0088,


		WM_NCMOUSEMOVE                 = 0x00A0,
		WM_NCLBUTTONDOWN               = 0x00A1,
		WM_NCLBUTTONUP                 = 0x00A2,
		WM_NCLBUTTONDBLCLK             = 0x00A3,
		WM_NCRBUTTONDOWN               = 0x00A4,
		WM_NCRBUTTONUP                 = 0x00A5,
		WM_NCRBUTTONDBLCLK             = 0x00A6,
		WM_NCMBUTTONDOWN               = 0x00A7,
		WM_NCMBUTTONUP                 = 0x00A8,
		WM_NCMBUTTONDBLCLK             = 0x00A9,
		WM_NCXBUTTONDOWN               = 0x00AB,
		WM_NCXBUTTONUP                 = 0x00AC,
		WM_NCXBUTTONDBLCLK             = 0x00AD,

		WM_INPUT_DEVICE_CHANGE         = 0x00FE,
		WM_INPUT                       = 0x00FF,

		WM_KEYFIRST                    = 0x0100,
		WM_KEYDOWN                     = 0x0100,
		WM_KEYUP                       = 0x0101,
		WM_CHAR                        = 0x0102,
		WM_DEADCHAR                    = 0x0103,
		WM_SYSKEYDOWN                  = 0x0104,
		WM_SYSKEYUP                    = 0x0105,
		WM_SYSCHAR                     = 0x0106,
		WM_SYSDEADCHAR                 = 0x0107,
		WM_UNICHAR                     = 0x0109,
		WM_KEYLAST                     = 0x0109,

		WM_IME_STARTCOMPOSITION        = 0x010D,
		WM_IME_ENDCOMPOSITION          = 0x010E,
		WM_IME_COMPOSITION             = 0x010F,
		WM_IME_KEYLAST                 = 0x010F,

		WM_INITDIALOG                  = 0x0110,
		WM_COMMAND                     = 0x0111,
		WM_SYSCOMMAND                  = 0x0112,
		WM_TIMER                       = 0x0113,
		WM_HSCROLL                     = 0x0114,
		WM_VSCROLL                     = 0x0115,
		WM_INITMENU                    = 0x0116,
		WM_INITMENUPOPUP               = 0x0117,
		WM_MENUSELECT                  = 0x011F,
		WM_MENUCHAR                    = 0x0120,
		WM_ENTERIDLE                   = 0x0121,
		WM_MENURBUTTONUP               = 0x0122,
		WM_MENUDRAG                    = 0x0123,
		WM_MENUGETOBJECT               = 0x0124,
		WM_UNINITMENUPOPUP             = 0x0125,
		WM_MENUCOMMAND                 = 0x0126,

		WM_CHANGEUISTATE               = 0x0127,
		WM_UPDATEUISTATE               = 0x0128,
		WM_QUERYUISTATE                = 0x0129,

		WM_CTLCOLORMSGBOX              = 0x0132,
		WM_CTLCOLOREDIT                = 0x0133,
		WM_CTLCOLORLISTBOX             = 0x0134,
		WM_CTLCOLORBTN                 = 0x0135,
		WM_CTLCOLORDLG                 = 0x0136,
		WM_CTLCOLORSCROLLBAR           = 0x0137,
		WM_CTLCOLORSTATIC              = 0x0138,
		MN_GETHMENU                    = 0x01E1,

		WM_MOUSEFIRST                  = 0x0200,
		WM_MOUSEMOVE                   = 0x0200,
		WM_LBUTTONDOWN                 = 0x0201,
		WM_LBUTTONUP                   = 0x0202,
		WM_LBUTTONDBLCLK               = 0x0203,
		WM_RBUTTONDOWN                 = 0x0204,
		WM_RBUTTONUP                   = 0x0205,
		WM_RBUTTONDBLCLK               = 0x0206,
		WM_MBUTTONDOWN                 = 0x0207,
		WM_MBUTTONUP                   = 0x0208,
		WM_MBUTTONDBLCLK               = 0x0209,
		WM_MOUSEWHEEL                  = 0x020A,
		WM_XBUTTONDOWN                 = 0x020B,
		WM_XBUTTONUP                   = 0x020C,
		WM_XBUTTONDBLCLK               = 0x020D,
		WM_MOUSEHWHEEL                 = 0x020E,

		WM_PARENTNOTIFY                = 0x0210,
		WM_ENTERMENULOOP               = 0x0211,
		WM_EXITMENULOOP                = 0x0212,

		WM_NEXTMENU                    = 0x0213,
		WM_SIZING                      = 0x0214,
		WM_CAPTURECHANGED              = 0x0215,
		WM_MOVING                      = 0x0216,

		WM_POWERBROADCAST              = 0x0218,

		WM_DEVICECHANGE                = 0x0219,

		WM_MDICREATE                   = 0x0220,
		WM_MDIDESTROY                  = 0x0221,
		WM_MDIACTIVATE                 = 0x0222,
		WM_MDIRESTORE                  = 0x0223,
		WM_MDINEXT                     = 0x0224,
		WM_MDIMAXIMIZE                 = 0x0225,
		WM_MDITILE                     = 0x0226,
		WM_MDICASCADE                  = 0x0227,
		WM_MDIICONARRANGE              = 0x0228,
		WM_MDIGETACTIVE                = 0x0229,


		WM_MDISETMENU                  = 0x0230,
		WM_ENTERSIZEMOVE               = 0x0231,
		WM_EXITSIZEMOVE                = 0x0232,
		WM_DROPFILES                   = 0x0233,
		WM_MDIREFRESHMENU              = 0x0234,

		WM_IME_SETCONTEXT              = 0x0281,
		WM_IME_NOTIFY                  = 0x0282,
		WM_IME_CONTROL                 = 0x0283,
		WM_IME_COMPOSITIONFULL         = 0x0284,
		WM_IME_SELECT                  = 0x0285,
		WM_IME_CHAR                    = 0x0286,
		WM_IME_REQUEST                 = 0x0288,
		WM_IME_KEYDOWN                 = 0x0290,
		WM_IME_KEYUP                   = 0x0291,

		WM_MOUSEHOVER                  = 0x02A1,
		WM_MOUSELEAVE                  = 0x02A3,
		WM_NCMOUSEHOVER                = 0x02A0,
		WM_NCMOUSELEAVE                = 0x02A2,

		WM_WTSSESSION_CHANGE           = 0x02B1,

		WM_TABLET_FIRST                = 0x02c0,
		WM_TABLET_LAST                 = 0x02df,

		WM_CUT                         = 0x0300,
		WM_COPY                        = 0x0301,
		WM_PASTE                       = 0x0302,
		WM_CLEAR                       = 0x0303,
		WM_UNDO                        = 0x0304,
		WM_RENDERFORMAT                = 0x0305,
		WM_RENDERALLFORMATS            = 0x0306,
		WM_DESTROYCLIPBOARD            = 0x0307,
		WM_DRAWCLIPBOARD               = 0x0308,
		WM_PAINTCLIPBOARD              = 0x0309,
		WM_VSCROLLCLIPBOARD            = 0x030A,
		WM_SIZECLIPBOARD               = 0x030B,
		WM_ASKCBFORMATNAME             = 0x030C,
		WM_CHANGECBCHAIN               = 0x030D,
		WM_HSCROLLCLIPBOARD            = 0x030E,
		WM_QUERYNEWPALETTE             = 0x030F,
		WM_PALETTEISCHANGING           = 0x0310,
		WM_PALETTECHANGED              = 0x0311,
		WM_HOTKEY                      = 0x0312,

		WM_PRINT                       = 0x0317,
		WM_PRINTCLIENT                 = 0x0318,

		WM_APPCOMMAND                  = 0x0319,

		WM_THEMECHANGED                = 0x031A,

		WM_CLIPBOARDUPDATE             = 0x031D,

		WM_DWMCOMPOSITIONCHANGED       = 0x031E,
		WM_DWMNCRENDERINGCHANGED       = 0x031F,
		WM_DWMCOLORIZATIONCOLORCHANGED = 0x0320,
		WM_DWMWINDOWMAXIMIZEDCHANGE    = 0x0321,

		WM_GETTITLEBARINFOEX           = 0x033F,

		WM_HANDHELDFIRST               = 0x0358,
		WM_HANDHELDLAST                = 0x035F,

		WM_AFXFIRST                    = 0x0360,
		WM_AFXLAST                     = 0x037F,

		WM_PENWINFIRST                 = 0x0380,
		WM_PENWINLAST                  = 0x038F,

		WM_APP                         = 0x8000,

		WM_USER                        = 0x0400,

		WM_REFLECT                     = 0x2000
	} WM_CONSTANTS;

	typedef enum WINDOWS_STYLES {
		WS_OVERLAPPED = 0x00000000,
		WS_POPUP = 0x80000000,
		WS_CHILD = 0x40000000,
		WS_MINIMIZE = 0x20000000,
		WS_VISIBLE = 0x10000000,
		WS_DISABLED = 0x08000000,
		WS_CLIPSIBLINGS = 0x04000000,
		WS_CLIPCHILDREN = 0x02000000,
		WS_MAXIMIZE = 0x01000000,
		WS_CAPTION = 0x00C00000,     /* WS_BORDER | WS_DLGFRAME  */
		WS_BORDER = 0x00800000,
		WS_DLGFRAME = 0x00400000,
		WS_VSCROLL = 0x00200000,
		WS_HSCROLL = 0x00100000,
		WS_SYSMENU = 0x00080000,
		WS_THICKFRAME = 0x00040000,
		WS_GROUP = 0x00020000,
		WS_TABSTOP = 0x00010000,
		WS_MINIMIZEBOX = 0x00020000,
		WS_MAXIMIZEBOX = 0x00010000,
		WS_TILED = 0x00000000,
		WS_ICONIC = 0x20000000,
		WS_SIZEBOX = 0x00040000,
		WS_OVERLAPPEDWINDOW = 0x00cf0000, 
		WS_POPUPWINDOW = 0x80880000,
		WS_CHILDWINDOW = 0x40000000,
		WS_TILEDWINDOW = 0x00cf0000
	} WINDOWS_STYLES;

	typedef enum WINDOWS_EXSTYLES {
		WS_EX_DLGMODALFRAME = 0x00000001,
		WS_EX_NOPARENTNOTIFY = 0x00000004,
		WS_EX_TOPMOST = 0x00000008,
		WS_EX_ACCEPTFILES = 0x00000010,
		WS_EX_TRANSPARENT = 0x00000020,
		WS_EX_MDICHILD = 0x00000040,
		WS_EX_TOOLWINDOW = 0x00000080,
		WS_EX_WINDOWEDGE = 0x00000100,
		WS_EX_CLIENTEDGE = 0x00000200,
		WS_EX_CONTEXTHELP = 0x00000400,
		WS_EX_RIGHT = 0x00001000,
		WS_EX_LEFT = 0x00000000,
		WS_EX_RTLREADING = 0x00002000,
		WS_EX_LTRREADING = 0x00000000,
		WS_EX_LEFTSCROLLBAR = 0x00004000,
		WS_EX_RIGHTSCROLLBAR = 0x00000000,
		WS_EX_CONTROLPARENT = 0x00010000,
		WS_EX_STATICEDGE = 0x00020000,
		WS_EX_APPWINDOW = 0x00040000,
		WS_EX_OVERLAPPEDWINDOW = 0x300,
		WS_EX_PALETTEWINDOW = 0x188,
		WS_EX_LAYERED = 0x00080000,
		WS_EX_NOINHERITLAYOUT = 0x00100000, // Disable inheritence of mirroring by children
		WS_EX_LAYOUTRTL = 0x00400000, // Right to left mirroring
		WS_EX_COMPOSITED = 0x02000000,
		WS_EX_NOACTIVATE = 0x08000000
	} WINDOWS_EXSTYLES;

	typedef struct MINMAXINFO {
		POINT ptReserved;
		POINT ptMaxSize;
		POINT ptMaxPosition;
		POINT ptMinTrackSize;
		POINT ptMaxTrackSize;
	} MINMAXINFO;

	[dllname("USER32.DLL")]
	module user32 {
		
		const int RT_RCDATA = 10;
		const int RT_FONT = 8;

		[entry("SendMessageW")]
		long SendMessage(
		  [in] long hWnd,
		  [in] long Msg,
		  [in] long wParam,
		  [in] void *lParam
		);

		[entry("CreateWindowExW")]
		long CreateWindowEx(
			[in] long dwExStyle,
			[in] LPWSTR lpClassName,
			[in] LPWSTR lpWindowName,
			[in] long dwStyle,
			[in] long X,
			[in] long Y,
			[in] long nWidth,
			[in] long nHeight,
			[in] long hWndParent,
			[in] long hMenu,
			[in] long hInstance,
			[in] void* lpParam
		);

		[entry("DestroyWindow")]
		long DestroyWindow(
			[in] long hWnd
		);

		[entry("OffsetRect")]
		long OffsetRect(
			[in] RECT* lpRect,
			[in] long x,
			[in] long y
		);

		[entry("InvalidateRect")]
		long InvalidateRect(
			[in] long hwnd,
			[in] void* lpRect,
			[in] long bErase 
		);

		[entry("DrawTextW")]
		long DrawText(
			[in] long hDC,
			[in] LPWSTR lpStr,
			[in] long nCount,
			[in] RECT* lpRect,
			[in] DRAWTEXT_FORMATS wFormat
		);

		[entry("MoveWindow")]
		long MoveWindow(
			[in] long hwnd,
			[in] long x,
			[in] long y,
			[in] long nWidth,
			[in] long nHeight,
			[in] long bRepaint
		);

		[entry("PtInRect")]
		long PtInRect(
			[in] RECT* lpRect,
			[in] long x,
			[in] long y
		);

		[entry("SetRect")]
		long SetRect(
			[in] RECT* lpRect,
			[in] long x1,
			[in] long y1,
			[in] long x2,
			[in] long y2
		);

		[entry("SetFocus")]
		long SetFocusAPI(
			[in] long hwnd
		);

		[entry("SetWindowTextW")]
		long SetWindowText(
			[in] long hWnd,
			[in] LPWSTR lpString
		);

		[entry("GetWindowTextW")]
		long GetWindowText(
			[in] long hWnd,
			[in] LPWSTR lpString,
			[in] long cch
		);

		[entry("GetWindowTextLengthW")]
		long GetWindowTextLength(
			[in] long hWnd
		);

		[entry("OpenClipboard")]
		long OpenClipboard(
			[in] long hWndNewOwner
		);

		[entry("EmptyClipboard")]
		long EmptyClipboard();

		[entry("CloseClipboard")]
		long CloseClipboard();

		[entry("SetClipboardData")]
		long SetClipboardData(
			[in] CLIPBOARD_FORMAT_CONSTS uFormat,
			[in] long hMem
		);

		[entry("GetClipboardData")]
		long GetClipboardData(
			[in] CLIPBOARD_FORMAT_CONSTS uFormat
		);

		[entry("IsClipboardFormatAvailable")]
		long IsClipboardFormatAvailable(
			[in] CLIPBOARD_FORMAT_CONSTS uFormat
		);

	}
