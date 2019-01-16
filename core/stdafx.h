#pragma once

#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#include <Windows.h>

// Common CRT headers
#include <list>
#include <algorithm>
#include <locale>
#include <sstream>
#include <iterator>
#include <functional>

// All the MAPI headers
#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIForm.h>
#include <MAPIWz.h>
#include <MAPIHook.h>
#include <MSPST.h>

#include <EdkMdb.h>
#include <ExchForm.h>
#include <EMSABTAG.H>
#include <IMessage.h>
#include <EdkGuid.h>
#include <TNEF.h>
#include <MAPIAux.h>

#include <core/res/Resource.h> // main symbols