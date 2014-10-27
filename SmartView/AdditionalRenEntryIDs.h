#pragma once
#include "SmartViewParser.h"

struct PersistElement
{
	WORD wElementID;
	WORD wElementDataSize;
	LPBYTE lpbElementData;
};

struct PersistData
{
	WORD wPersistID;
	WORD wDataElementsSize;
	WORD wDataElementCount;
	PersistElement* ppeDataElement;

	size_t JunkDataSize;
	LPBYTE JunkData;
};

class AdditionalRenEntryIDs : public SmartViewParser
{
public:
	AdditionalRenEntryIDs(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~AdditionalRenEntryIDs();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();
	void BinToPersistData(_Out_ PersistData* ppdPersistData);

	WORD m_wPersistDataCount;
	PersistData* m_ppdPersistData;
};