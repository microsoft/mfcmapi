#include "stdafx.h"
#include "CommentData.h"

CommentData::CommentData(_In_ LPSPropValue lpOldProp)
{
	m_lpOldProp = lpOldProp;
	m_lpNewProp = nullptr;
}