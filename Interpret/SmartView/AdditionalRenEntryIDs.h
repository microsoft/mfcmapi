#pragma once
#include "SmartViewParser.h"

struct PersistElement
{
	WORD wElementID;
	WORD wElementDataSize;
	vector<BYTE> lpbElementData;
};

struct PersistData
{
	WORD wPersistID;
	WORD wDataElementsSize;
	vector<PersistElement> ppeDataElement;

	vector<BYTE> JunkData;
};

class AdditionalRenEntryIDs : public SmartViewParser
{
private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;
	PersistData BinToPersistData();

	vector<PersistData> m_ppdPersistData;
};