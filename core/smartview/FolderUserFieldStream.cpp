#include <core/stdafx.h>
#include <core/smartview/FolderUserFieldStream.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void FolderFieldDefinition ::parse()
	{
		FieldType = blockT<DWORD>::parse(parser);
		FieldNameLength = blockT<WORD>::parse(parser);

		if (*FieldNameLength && *FieldNameLength < _MaxEntriesSmall)
		{
			if (unicode)
			{
				FieldNameW = blockStringW::parse(parser, *FieldNameLength);
			}
			else
			{
				FieldNameA = blockStringA::parse(parser, *FieldNameLength);
			}
		}

		PropSetGuid = blockT<GUID>::parse(parser);
		fcapm = blockT<DWORD>::parse(parser);
		dwString = blockT<DWORD>::parse(parser);
		dwBitmap = blockT<DWORD>::parse(parser);
		dwDisplay = blockT<DWORD>::parse(parser);
		iFmt = blockT<DWORD>::parse(parser);
		wszFormulaLength = blockT<WORD>::parse(parser);
		if (*wszFormulaLength && *wszFormulaLength < _MaxEntriesLarge)
		{
			wszFormula = blockStringW::parse(parser, *wszFormulaLength);
		}
	}

	void FolderFieldDefinition ::parseBlocks()
	{
		if (unicode)
		{
			setText(L"FolderFieldDefinitionW");
		}
		else
		{
			setText(L"FolderFieldDefinitionA");
		}

		auto szFieldType = flags::InterpretFlags(flagFolderType, *FieldType);
		addChild(FieldType, L"FieldType = 0x%1!08X! = %2!ws!", FieldType->getData(), szFieldType.c_str());
		addChild(FieldNameLength, L"FieldNameLength = 0x%1!08X! = %1!d!", FieldNameLength->getData());
		if (unicode)
		{
			addChild(FieldNameW, L"FieldName = %1!ws!", FieldNameW->c_str());
		}
		else
		{
			addChild(FieldNameA, L"FieldName = %1!hs!", FieldNameA->c_str());
		}

		auto szGUID = guid::GUIDToString(*PropSetGuid);
		addChild(PropSetGuid, L"PropSetGuid = %1!ws!", szGUID.c_str());
		auto szFieldcap = flags::InterpretFlags(flagFieldCap, *fcapm);
		addChild(fcapm, L"fcapm = 0x%1!08X! = %2!ws!", fcapm->getData(), szFieldcap.c_str());
		addChild(dwString, L"dwString = 0x%1!08X!", dwString->getData());
		addChild(dwBitmap, L"dwBitmap = 0x%1!08X!", dwBitmap->getData());
		addChild(dwDisplay, L"dwDisplay = 0x%1!08X!", dwDisplay->getData());
		addChild(iFmt, L"iFmt = 0x%1!08X!", iFmt->getData());
		addChild(wszFormulaLength, L"wszFormulaLength = 0x%1!04X! = %1!d!", wszFormulaLength->getData());
		addChild(wszFormula, L"wszFormula = %1!ws!", wszFormula->c_str());
	}

	void FolderUserFieldStream::parse()
	{
		m_FolderUserFieldsAnsiCount = blockT<DWORD>::parse(parser);
		if (*m_FolderUserFieldsAnsiCount && *m_FolderUserFieldsAnsiCount < _MaxEntriesSmall)
		{
			m_FieldDefinitionsA.reserve(*m_FolderUserFieldsAnsiCount);
			for (DWORD i = 0; i < *m_FolderUserFieldsAnsiCount; i++)
			{
				if (parser->empty()) continue;
				auto fd = std::make_shared<FolderFieldDefinition>(false);
				fd->block::parse(parser, 0, false);
				m_FieldDefinitionsA.emplace_back(fd);
			}
		}

		m_FolderUserFieldsUnicodeCount = blockT<DWORD>::parse(parser);
		if (*m_FolderUserFieldsUnicodeCount && *m_FolderUserFieldsUnicodeCount < _MaxEntriesSmall)
		{
			m_FieldDefinitionsW.reserve(*m_FolderUserFieldsUnicodeCount);
			for (DWORD i = 0; i < *m_FolderUserFieldsUnicodeCount; i++)
			{
				if (parser->empty()) continue;
				auto fd = std::make_shared<FolderFieldDefinition>(true);
				fd->block::parse(parser, 0, false);
				m_FieldDefinitionsW.emplace_back(fd);
			}
		}
	}

	void FolderUserFieldStream::parseBlocks()
	{
		setText(L"Folder User Field Stream");
		addChild(
			m_FolderUserFieldsAnsiCount,
			L"FolderUserFieldAnsi.FieldDefinitionCount = %1!d!",
			m_FolderUserFieldsAnsiCount->getData());

		// Add child nodes to m_FolderUserFieldsAnsiCount
		if (m_FolderUserFieldsAnsiCount && !m_FieldDefinitionsA.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsA)
			{
				m_FolderUserFieldsAnsiCount->addChild(fieldDefinition, L"Field %1!d!", i++);
			}
		}

		addChild(
			m_FolderUserFieldsUnicodeCount,
			L"FolderUserFieldUnicode.FieldDefinitionCount = %1!d!",
			m_FolderUserFieldsUnicodeCount->getData());

		// Add child nodes to m_FolderUserFieldsUnicodeCount
		if (m_FolderUserFieldsUnicodeCount && !m_FieldDefinitionsW.empty())
		{
			auto i = 0;
			for (const auto& fieldDefinition : m_FieldDefinitionsW)
			{
				m_FolderUserFieldsUnicodeCount->addChild(fieldDefinition, L"Field %1!d!", i++);
			}
		}
	}
} // namespace smartview