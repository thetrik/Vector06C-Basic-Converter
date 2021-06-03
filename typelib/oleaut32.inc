	
	typedef struct SAFEARRAYBOUND {
		long cElements;
		long lLBound;
	} SAFEARRAYBOUND;

	typedef struct SAFEARRAY {
		short cDims;
		short fFeatures;
		long cbElements;
		long cLocks;
		long pvData;
		//SAFEARRAYBOUND rgsabound[1];
	} SAFEARRAY;

