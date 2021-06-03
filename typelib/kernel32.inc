	typedef enum FILE_ATTRIBUTES_AND_FLAGS {

		FILE_ATTRIBUTE_READONLY = 0x00000001,  
		FILE_ATTRIBUTE_HIDDEN = 0x00000002,  
		FILE_ATTRIBUTE_SYSTEM = 0x00000004,  
		FILE_ATTRIBUTE_DIRECTORY = 0x00000010,  
		FILE_ATTRIBUTE_ARCHIVE = 0x00000020,  
		FILE_ATTRIBUTE_DEVICE = 0x00000040,  
		FILE_ATTRIBUTE_NORMAL = 0x00000080,  
		FILE_ATTRIBUTE_TEMPORARY = 0x00000100,  
		FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200,  
		FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400,  
		FILE_ATTRIBUTE_COMPRESSED = 0x00000800,  
		FILE_ATTRIBUTE_OFFLINE = 0x00001000,  
		FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000,  
		FILE_ATTRIBUTE_ENCRYPTED = 0x00004000,  
		FILE_ATTRIBUTE_VIRTUAL = 0x00010000,  
		FILE_FLAG_WRITE_THROUGH = 0x80000000,
		FILE_FLAG_OVERLAPPED = 0x40000000,
		FILE_FLAG_NO_BUFFERING = 0x20000000,
		FILE_FLAG_RANDOM_ACCESS = 0x10000000,
		FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000,
		FILE_FLAG_DELETE_ON_CLOSE = 0x04000000,
		FILE_FLAG_BACKUP_SEMANTICS = 0x02000000,
		FILE_FLAG_POSIX_SEMANTICS = 0x01000000,
		FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000,
		FILE_FLAG_OPEN_NO_RECALL = 0x00100000,
		FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000

	} FILE_ATTRIBUTES_AND_FLAGS;

	typedef enum FILE_SHARE {

		FILE_SHARE_DELETE = 4,
		FILE_SHARE_READ = 1,
		FILE_SHARE_WRITE = 2,
		FILE_SHARE_NONE = 0

	} FILE_SHARE;

	typedef enum PAGE_ACCESS {
		PAGE_EXECUTE_READ=0x20,
		PAGE_EXECUTE_READWRITE=0x40,
		PAGE_EXECUTE_WRITECOPY=0x80,
		PAGE_READONLY=0x02,
		PAGE_READWRITE=0x04,
		PAGE_WRITECOPY=0x08,
		PAGE_NOCACHE = 0x200,
		PAGE_NOACCESS = 1,
		PAGE_GUARD = 0x100,
		SEC_COMMIT=0x8000000,
		SEC_IMAGE=0x1000000,
		SEC_IMAGE_NO_EXECUTE=0x11000000,
		SEC_LARGE_PAGES=0x80000000,
		SEC_NOCACHE=0x10000000,
		SEC_RESERVE=0x4000000,
		SEC_WRITECOMBINE=0x40000000
	} PAGE_ACCESS;

	typedef enum CREATION_DISPOSITION {
		 CREATE_ALWAYS = 2,
		 CREATE_NEW = 1,
		 OPEN_ALWAYS = 4,
		 OPEN_EXISTING = 3,
		 TRUNCATE_EXISTING = 5
	} CREATION_DISPOSITION;

	typedef enum FILE_MAP {
		FILE_MAP_READ = 4,
		FILE_MAP_WRITE = 2,
		FILE_MAP_COPY = 1,
		FILE_MAP_ALL_ACCESS = 0xF001F,
		FILE_MAP_EXECUTE = 0x20
	} FILE_MAP;

	typedef enum CODEPAGES {
		CP_WINANSI = 1004,    /* default codepage for windows & old DDE convs. */
		CP_WINUNICODE = 1200,
		CP_ACP = 0,           // default to ANSI code page
		CP_OEMCP = 1,           // default to OEM  code page
		CP_MACCP = 2,           // default to MAC  code page
		CP_THREAD_ACP = 3,           // current thread's ANSI code page
		CP_SYMBOL = 42,          // SYMBOL translations
		CP_UTF7 = 65000,       // UTF-7 translation
		CP_UTF8 = 65001       // UTF-8 translation
	} CODEPAGES;

	typedef struct LARGE_INTEGER {
		long LowPart;
		long HighPart;
	} LARGE_INTEGER;

	[dllname("KERNEL32.DLL")]
	module kernel32 {

		const int MAX_PATH = 260;

		const int GENERIC_READ = 0x80000000;
		const int GENERIC_WRITE = 0x40000000;
		const int GENERIC_EXECUTE = 0x20000000;
		const int GENERIC_ALL = 0x10000000;

		const int INVALID_HANDLE_VALUE = -1;

		[entry("lstrlenW")]
		int lstrlenW(
			[in] void* lpString
		);

		[entry("lstrcpynA")]
		int lstrcpynA(
			[in] void* lpString1, 
			[in] void* lpString2,
			[in] long nMaxLen
		);

		[entry("lstrcpynW")]
		int lstrcpynW(
			[in] void* lpString1, 
			[in] void* lpString2,
			[in] long nMaxLen
		);

		[entry("GetFileSizeEx")]
		long GetFileSizeEx(
			[in] long hFile,
			[in, out] LARGE_INTEGER *lpFileSize
		);

		[entry("RtlMoveMemory")]
		void memcpy(
			[in] void* Destination,
			[in] void* Source,
			[in] long Length
		);

		[entry("CloseHandle")]
		long CloseHandle(
			[in] long hObject);

		[entry("CreateFileW")]
		long CreateFile(
			[in] LPWSTR lpFileName,
			[in] long dwDesiredAccess,
			[in] FILE_SHARE dwShareMode,
			[in] void *lpSecurityAttributes,
			[in] CREATION_DISPOSITION dwCreationDisposition,
			[in] FILE_ATTRIBUTES_AND_FLAGS dwFlagsAndAttributes,
			[in] long hTemplateFile
		);

		[entry("ReadFile")]
		long ReadFile(
			[in] long hFile,
			[in] void *lpBuffer,
			[in] long nNumberOfBytesToRead,
			[in] long *lpNumberOfBytesRead,
			[in] void *lpOverlapped
		);

		[entry("WriteFile")]
		long WriteFile(
			[in] long hFile,
			[in] void *lpBuffer,
			[in] long nNumberOfBytesToWrite,
			[in] long *lpNumberOfBytesWritten,
			[in] void *lpOverlapped
		);

		[entry("WideCharToMultiByte")]
		long WideCharToMultiByte(
			[in] long CodePage,
			[in] long dwFlags,
			[in] void* lpWideCharStr,
			[in] long cchWideChar,
			[in] void* lpMultiByteStr,
			[in] long cchMultiByte,
			[in] void* lpDefaultChar,
			[in] long* lpUsedDefaultChar
		);

		[entry("MultiByteToWideChar")]
		long MultiByteToWideChar(
			[in] long CodePage,
			[in] long dwFlags,
			[in] void* lpMultiByteStr,
			[in] long cchMultiByte,
			[in] void* lpWideCharStr,
			[in] long cchWideChar
		);

		[entry("CreateFileMappingW")]
		long CreateFileMapping(
			[in] long hFile,
			[in] void* lpFileMappigAttributes,
			[in] PAGE_ACCESS flProtect,
			[in] long dwMaximumSizeHigh,
			[in] long dwMaximumSizeLow,
			[in] LPWSTR lpName
		);

		[entry("MapViewOfFile")]
		long MapViewOfFile(
			[in] long hFileMappingObject,
			[in] FILE_MAP dwDesiredAccess,
			[in] long dwFileOffsetHigh,
			[in] long dwFileOffsetLow,
			[in] long dwNumberOfBytesToMap
		);

		[entry("UnmapViewOfFile")]
		long UnmapViewOfFile(
			[in] void* lpBaseAddress
		);

		[entry("WritePrivateProfileStringW")]
		long WritePrivateProfileString(
			[in] LPWSTR lpAppName,
			[in] LPWSTR lpKeyName,
			[in] LPWSTR lpString,
			[in] LPWSTR lpFileName
		);

		[entry("GetPrivateProfileStringW")]
		long GetPrivateProfileString(
			[in] LPWSTR lpAppName,
			[in] LPWSTR lpKeyName,
			[in] LPWSTR lpDefault,
			[in] LPWSTR lpReturnedString,
			[in] long nSize,
			[in] LPWSTR lpFileName
		);

	}
