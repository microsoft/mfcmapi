#pragma once

enum class modifyType
{
	REQUEST_MODIFY,
	DO_NOT_REQUEST_MODIFY
};

enum createDialogType
{
	CALL_CREATE_DIALOG,
	DO_NOT_CALL_CREATE_DIALOG
};

enum statusPaneType
{
	DATA1,
	DATA2,
	INFOTEXT,
	NUMPANES
};

// Flags to indicate which contents/hierarchy tables to render
enum _mfcmapiDisplayFlagsEnum
{
	dfNormal = 0x0000,
	dfAssoc = 0x0001,
	dfDeleted = 0x0002,
};

enum __mfcmapiRestrictionTypeEnum
{
	mfcmapiNO_RESTRICTION,
	mfcmapiNORMAL_RESTRICTION,
	mfcmapiFINDROW_RESTRICTION
};
