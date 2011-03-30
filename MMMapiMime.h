#pragma once
// MMMapiMime.h : MAPI <-> MIME conversion for MrMAPI

// Flags to control conversion
#define MAPIMIME_TOMAPI      0x00000001
#define MAPIMIME_TOMIME      0x00000002
#define MAPIMIME_RFC822      0x00000004
#define MAPIMIME_WRAP        0x00000008
#define MAPIMIME_ENCODING    0x00000010
#define MAPIMIME_ADDRESSBOOK 0x00000020
#define MAPIMIME_UNICODE     0x00000040
#define MAPIMIME_CHARSET     0x00000080

void DoMAPIMIME(_In_ MYOPTIONS ProgOpts);