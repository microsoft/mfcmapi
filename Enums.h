#pragma once

enum __mfcmapiModifyEnum
{
	mfcmapiREQUEST_MODIFY,
	mfcmapiDO_NOT_REQUEST_MODIFY
};

enum __mfcmapiCreateDialogEnum
{
	mfcmapiCALL_CREATE_DIALOG,
	mfcmapiDO_NOT_CALL_CREATE_DIALOG
};

enum __StatusPaneEnum
{
	STATUSDATA1,
	STATUSDATA2,
	STATUSINFOTEXT,
	STATUSBARNUMPANES
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
