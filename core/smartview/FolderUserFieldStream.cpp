#include <core/stdafx.h>
#include <core/smartview/FolderUserFieldStream.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	FolderFieldDefinitionA ::FolderFieldDefinitionA(const std::shared_ptr<binaryParser>& parser)
	{
		FieldType = blockT<DWORD>::parse(parser);
		FieldNameLength = blockT<WORD>::parse(parser);

		if (*FieldNameLength && *FieldNameLength < _MaxEntriesSmall)
		{
			FieldName = blockStringA::parse(parser, *FieldNameLength);
		}

		Common.parse(parser);
	}

	FolderFieldDefinitionW ::FolderFieldDefinitionW(const std::shared_ptr<binaryParser>& parser)
	{
		FieldType = blockT<DWORD>::parse(parser);
		FieldNameLength = blockT<WORD>::parse(parser);

		if (*FieldNameLength && *FieldNameLength < _MaxEntriesSmall)
		{
			FieldName = blockStringW::parse(parser, *FieldNameLength);
		}

		Common.parse(parser);
	}

	void FolderUserFieldStream::parse()
	{
		m_FolderUserFieldsAnsiCount = blockT<DWORD>::parse(m_Parser);
		if (*m_FolderUserFieldsAnsiCount && *m_FolderUserFieldsAnsiCount < _MaxEntriesSmall)
		{
			m_FieldDefinitionsA.reserve(*m_FolderUserFieldsAnsiCount);
			for (DWORD i = 0; i < *m_FolderUserFieldsAnsiCount; i++)
			{
				if (m_Parser->empty()) continue;
				m_FieldDefinitionsA.emplace_back(std::make_shared<FolderFieldDefinitionA>(m_Parser));
			}
		}

		m_FolderUserFieldsUnicodeCount = blockT<DWORD>::parse(m_Parser);
		if (*m_FolderUserFieldsUnicodeCount && *m_FolderUserFieldsUnicodeCount < _MaxEntriesSmall)
		{
			m_FieldDefinitionsW.reserve(*m_FolderUserFieldsUnicodeCount);
			for (DWORD i = 0; i < *m_FolderUserFieldsUnicodeCount; i++)
			{
				if (m_Parser->empty()) continue;
				m_FieldDefinitionsW.emplace_back(std::make_shared<FolderFieldDefinitionW>(m_Parser));
			}
		}
	}

	void FolderFieldDefinitionCommon::parse(const std::shared_ptr<binaryParser>& parser)
	{
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

	void FolderUserFieldStream::parseBlocks()
	{
		setRoot(L"Folder User Field Stream\r\n");
		addChild(
			m_FolderUserFieldsAnsiCount,
			L"FolderUserFieldAnsi.FieldDefinitionCount = %1!d!\r\n",
			m_FolderUserFieldsAnsiCount->getData());

		// Add child nodes to m_FolderUserFieldsAnsiCount
		if (m_FolderUserFieldsAnsiCount && !m_FieldDefinitionsA.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsA)
			{
				m_FolderUserFieldsAnsiCount->terminateBlock();
				m_FolderUserFieldsAnsiCount->addBlankLine();

				auto fieldBlock = std::make_shared<block>();
				fieldBlock->setText(L"Field %1!d!\r\n", i++);
				m_FolderUserFieldsAnsiCount->addChild(fieldBlock);

				auto szFieldType = flags::InterpretFlags(flagFolderType, *fieldDefinition->FieldType);
				fieldBlock->addChild(
					fieldDefinition->FieldType,
					L"FieldType = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition->FieldType->getData(),
					szFieldType.c_str());
				fieldBlock->addChild(
					fieldDefinition->FieldNameLength,
					L"FieldNameLength = 0x%1!08X! = %1!d!\r\n",
					fieldDefinition->FieldNameLength->getData());
				fieldBlock->addChild(
					fieldDefinition->FieldName, L"FieldName = %1!hs!\r\n", fieldDefinition->FieldName->c_str());

				auto szGUID = guid::GUIDToString(*fieldDefinition->Common.PropSetGuid);
				fieldBlock->addChild(fieldDefinition->Common.PropSetGuid, L"PropSetGuid = %1!ws!\r\n", szGUID.c_str());
				auto szFieldcap = flags::InterpretFlags(flagFieldCap, *fieldDefinition->Common.fcapm);
				fieldBlock->addChild(
					fieldDefinition->Common.fcapm,
					L"fcapm = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition->Common.fcapm->getData(),
					szFieldcap.c_str());
				fieldBlock->addChild(
					fieldDefinition->Common.dwString,
					L"dwString = 0x%1!08X!\r\n",
					fieldDefinition->Common.dwString->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.dwBitmap,
					L"dwBitmap = 0x%1!08X!\r\n",
					fieldDefinition->Common.dwBitmap->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.dwDisplay,
					L"dwDisplay = 0x%1!08X!\r\n",
					fieldDefinition->Common.dwDisplay->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.iFmt, L"iFmt = 0x%1!08X!\r\n", fieldDefinition->Common.iFmt->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.wszFormulaLength,
					L"wszFormulaLength = 0x%1!04X! = %1!d!\r\n",
					fieldDefinition->Common.wszFormulaLength->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.wszFormula,
					L"wszFormula = %1!ws!",
					fieldDefinition->Common.wszFormula->c_str());
			}
		}

		addBlankLine();
		addChild(
			m_FolderUserFieldsUnicodeCount,
			L"FolderUserFieldUnicode.FieldDefinitionCount = %1!d!\r\n",
			m_FolderUserFieldsUnicodeCount->getData());

		// Add child nodes to m_FolderUserFieldsUnicodeCount
		if (m_FolderUserFieldsUnicodeCount && !m_FieldDefinitionsW.empty())
		{
			auto i = 0;
			for (const auto& fieldDefinition : m_FieldDefinitionsW)
			{
				m_FolderUserFieldsUnicodeCount->terminateBlock();
				m_FolderUserFieldsUnicodeCount->addBlankLine();

				auto fieldBlock = std::make_shared<block>();
				fieldBlock->setText(L"Field %1!d!\r\n", i++);
				m_FolderUserFieldsUnicodeCount->addChild(fieldBlock);

				auto szFieldType = flags::InterpretFlags(flagFolderType, *fieldDefinition->FieldType);
				fieldBlock->addChild(
					fieldDefinition->FieldType,
					L"FieldType = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition->FieldType->getData(),
					szFieldType.c_str());
				fieldBlock->addChild(
					fieldDefinition->FieldNameLength,
					L"FieldNameLength = 0x%1!08X! = %1!d!\r\n",
					fieldDefinition->FieldNameLength->getData());
				fieldBlock->addChild(
					fieldDefinition->FieldName, L"FieldName = %1!ws!\r\n", fieldDefinition->FieldName->c_str());

				auto szGUID = guid::GUIDToString(*fieldDefinition->Common.PropSetGuid);
				fieldBlock->addChild(fieldDefinition->Common.PropSetGuid, L"PropSetGuid = %1!ws!\r\n", szGUID.c_str());
				auto szFieldcap = flags::InterpretFlags(flagFieldCap, *fieldDefinition->Common.fcapm);
				fieldBlock->addChild(
					fieldDefinition->Common.fcapm,
					L"fcapm = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition->Common.fcapm->getData(),
					szFieldcap.c_str());
				fieldBlock->addChild(
					fieldDefinition->Common.dwString,
					L"dwString = 0x%1!08X!\r\n",
					fieldDefinition->Common.dwString->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.dwBitmap,
					L"dwBitmap = 0x%1!08X!\r\n",
					fieldDefinition->Common.dwBitmap->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.dwDisplay,
					L"dwDisplay = 0x%1!08X!\r\n",
					fieldDefinition->Common.dwDisplay->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.iFmt, L"iFmt = 0x%1!08X!\r\n", fieldDefinition->Common.iFmt->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.wszFormulaLength,
					L"wszFormulaLength = 0x%1!04X! = %1!d!\r\n",
					fieldDefinition->Common.wszFormulaLength->getData());
				fieldBlock->addChild(
					fieldDefinition->Common.wszFormula,
					L"wszFormula = %1!ws!",
					fieldDefinition->Common.wszFormula->c_str());
			}
		}
	}
} // namespace smartview