	typedef struct LOGFONT {
		long lfHeight;
		long lfWidth;
		long lfEscapement;
		long lfOrientation;
		long lfWeight;
		unsigned char lfItalic;
		unsigned char lfUnderline;
		unsigned char lfStrikeOut;
		unsigned char lfCharSet;
		unsigned char lfOutPrecision;
		unsigned char lfClipPrecision;
		unsigned char lfQuality;
		unsigned char lfPitchAndFamily;
		short lfFaceName[/* LF_FACESIZE */ 32];
	} LOGFONT;
	
	typedef enum LOGFONT_QUALITY {
		DEFAULT_QUALITY = 0,
		DRAFT_QUALITY = 1,
		PROOF_QUALITY = 2,
		NONANTIALIASED_QUALITY = 3,
		ANTIALIASED_QUALITY = 4,
		CLEARTYPE_QUALITY = 5,
		CLEARTYPE_NATURAL_QUALITY = 6
	} LOGFONT_QUALITY;

	[dllname("GDI32.DLL")]
	module gdi32 {
		
		const int LF_FACESIZE = 32;

		[entry("CreateFontIndirectW")]
		long CreateFontIndirect(
			[in] LOGFONT* lpLogFont
		);

		[entry("DeleteObject")]
		long DeleteObject(
			[in] long hObject
		);

		[entry("GetObjectW")]
		long GetObjectAPI(
			[in] long hObject,
			[in] long nCount,
			[in] void* lpObject
		);

		[entry("AddFontMemResourceEx")]
		long AddFontMemResourceEx(
			[in] void* pFileView,
			[in] long cjSize,
			[in] long pvResrved,
			[in] long* pNumFonts
		);

		[entry("RemoveFontMemResourceEx")]
		long RemoveFontMemResourceEx(
			[in] long h
		);

	}
