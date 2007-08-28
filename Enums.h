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
	STATUSLEFTPANE,
	STATUSMIDDLEPANE,
	STATUSRIGHTPANE,
	STATUSBARNUMPANES
};

enum __mfcmapiDeletedItemsEnum
{
	mfcmapiSHOW_DELETED_ITEMS,
	mfcmapiDO_NOT_SHOW_DELETED_ITEMS
};

enum __mfcmapiAssociatedContentsEnum
{
	mfcmapiSHOW_NORMAL_CONTENTS,
	mfcmapiSHOW_ASSOC_CONTENTS
};

enum __mfcmapiRestrictionTypeEnum
{
	mfcmapiNO_RESTRICTION,
	mfcmapiNORMAL_RESTRICTION,
	mfcmapiFINDROW_RESTRICTION
};
