#include "stdafx.h"
#include "CommentData.h"

CommentData::CommentData(_In_ const _SPropValue* lpOldProp)
{
	m_lpOldProp = lpOldProp;
	m_lpNewProp = nullptr;
}