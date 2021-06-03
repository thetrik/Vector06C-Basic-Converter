
	[dllname("SHLWAPI.DLL")]
	module shlwapi {

		[entry("HashData")]
		long HashData(
			[in] void* pbData,
			[in] long cbData,
			[in] void *pbHash,
			[in] long cbHash
		);

		[entry("PathFindExtensionW")]
		long PathFindExtension(
			[in] void* pszPath
		);

		[entry("PathFindFileNameW")]
		long PathFindFileName(
			[in] LPWSTR pszPath
		);

		[entry("PathRemoveFileSpecW")]
		long PathRemoveFileSpec(
			[in] LPWSTR pszPath
		);

	}
