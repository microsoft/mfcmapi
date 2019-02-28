#pragma once
#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#define _WINSOCKAPI_ // stops windows.h including winsock.h
#include <Windows.h>
#include <tchar.h>

// Common CRT headers
#include <list>
#include <algorithm>
#include <locale>
#include <sstream>
#include <iterator>
#include <functional>
#include <map>
#include <unordered_map>
#include <utility>
#include <vector>
#include <cassert>
#include <stack>
#include <deque>

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
#include <UnitTest/res/resource.h> // unittest symbols