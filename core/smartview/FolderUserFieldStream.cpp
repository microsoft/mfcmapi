#include <core/stdafx.h>
#include <core/smartview/FolderUserFieldStream.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void FolderUserFieldStream::Parse()
	{
		m_FolderUserFieldsAnsiCount = m_Parser.Get<DWORD>();
		if (m_FolderUserFieldsAnsiCount && m_FolderUserFieldsAnsiCount < _MaxEntriesSmall)
		{
			m_FieldDefinitionsA.reserve(m_FolderUserFieldsAnsiCount);
			for (DWORD i = 0; i < m_FolderUserFieldsAnsiCount; i++)
			{
				if (m_Parser.empty()) continue;
				FolderFieldDefinitionA folderFieldDefinitionA;
				folderFieldDefinitionA.FieldType = m_Parser.Get<DWORD>();
				folderFieldDefinitionA.FieldNameLength = m_Parser.Get<WORD>();

				if (folderFieldDefinitionA.FieldNameLength && folderFieldDefinitionA.FieldNameLength < _MaxEntriesSmall)
				{
					folderFieldDefinitionA.FieldName = m_Parser.GetStringA(folderFieldDefinitionA.FieldNameLength);
				}

				folderFieldDefinitionA.Common = BinToFolderFieldDefinitionCommon();
				m_FieldDefinitionsA.push_back(folderFieldDefinitionA);
			}
		}

		m_FolderUserFieldsUnicodeCount = m_Parser.Get<DWORD>();
		if (m_FolderUserFieldsUnicodeCount && m_FolderUserFieldsUnicodeCount < _MaxEntriesSmall)
		{
			m_FieldDefinitionsW.reserve(m_FolderUserFieldsUnicodeCount);
			for (DWORD i = 0; i < m_FolderUserFieldsUnicodeCount; i++)
			{
				if (m_Parser.empty()) continue;
				FolderFieldDefinitionW folderFieldDefinitionW;
				folderFieldDefinitionW.FieldType = m_Parser.Get<DWORD>();
				folderFieldDefinitionW.FieldNameLength = m_Parser.Get<WORD>();

				if (folderFieldDefinitionW.FieldNameLength && folderFieldDefinitionW.FieldNameLength < _MaxEntriesSmall)
				{
					folderFieldDefinitionW.FieldName = m_Parser.GetStringW(folderFieldDefinitionW.FieldNameLength);
				}

				folderFieldDefinitionW.Common = BinToFolderFieldDefinitionCommon();
				m_FieldDefinitionsW.push_back(folderFieldDefinitionW);
			}
		}
	}

	FolderFieldDefinitionCommon FolderUserFieldStream::BinToFolderFieldDefinitionCommon()
	{
		FolderFieldDefinitionCommon common;

		common.PropSetGuid = m_Parser.Get<GUID>();
		common.fcapm = m_Parser.Get<DWORD>();
		common.dwString = m_Parser.Get<DWORD>();
		common.dwBitmap = m_Parser.Get<DWORD>();
		common.dwDisplay = m_Parser.Get<DWORD>();
		common.iFmt = m_Parser.Get<DWORD>();
		common.wszFormulaLength = m_Parser.Get<WORD>();
		if (common.wszFormulaLength && common.wszFormulaLength < _MaxEntriesLarge)
		{
			common.wszFormula = m_Parser.GetStringW(common.wszFormulaLength);
		}

		return common;
	}

	void FolderUserFieldStream::ParseBlocks()
	{
		setRoot(L"Folder User Field Stream\r\n");

		// Add child nodes to m_FolderUserFieldsAnsiCount before adding it to our output
		if (m_FolderUserFieldsAnsiCount && !m_FieldDefinitionsA.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsA)
			{
				auto fieldBlock = block{};
				fieldBlock.setText(L"Field %1!d!\r\n", i++);

				auto szFieldType = flags::InterpretFlags(flagFolderType, fieldDefinition.FieldType);
				fieldBlock.addBlock(
					fieldDefinition.FieldType,
					L"FieldType = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition.FieldType.getData(),
					szFieldType.c_str());
				fieldBlock.addBlock(
					fieldDefinition.FieldNameLength,
					L"FieldNameLength = 0x%1!08X! = %1!d!\r\n",
					fieldDefinition.FieldNameLength.getData());
				fieldBlock.addBlock(
					fieldDefinition.FieldName, L"FieldName = %1!hs!\r\n", fieldDefinition.FieldName.c_str());

				auto szGUID = guid::GUIDToString(fieldDefinition.Common.PropSetGuid);
				fieldBlock.addBlock(fieldDefinition.Common.PropSetGuid, L"PropSetGuid = %1!ws!\r\n", szGUID.c_str());
				auto szFieldcap = flags::InterpretFlags(flagFieldCap, fieldDefinition.Common.fcapm);
				fieldBlock.addBlock(
					fieldDefinition.Common.fcapm,
					L"fcapm = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition.Common.fcapm.getData(),
					szFieldcap.c_str());
				fieldBlock.addBlock(
					fieldDefinition.Common.dwString,
					L"dwString = 0x%1!08X!\r\n",
					fieldDefinition.Common.dwString.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.dwBitmap,
					L"dwBitmap = 0x%1!08X!\r\n",
					fieldDefinition.Common.dwBitmap.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.dwDisplay,
					L"dwDisplay = 0x%1!08X!\r\n",
					fieldDefinition.Common.dwDisplay.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.iFmt, L"iFmt = 0x%1!08X!\r\n", fieldDefinition.Common.iFmt.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.wszFormulaLength,
					L"wszFormulaLength = 0x%1!04X! = %1!d!\r\n",
					fieldDefinition.Common.wszFormulaLength.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.wszFormula,
					L"wszFormula = %1!ws!",
					fieldDefinition.Common.wszFormula.c_str());

				m_FolderUserFieldsAnsiCount.terminateBlock();
				m_FolderUserFieldsAnsiCount.addBlankLine();
				m_FolderUserFieldsAnsiCount.addBlock(fieldBlock);
			}
		}

		m_FolderUserFieldsAnsiCount.terminateBlock();
		addBlock(
			m_FolderUserFieldsAnsiCount,
			L"FolderUserFieldAnsi.FieldDefinitionCount = %1!d!\r\n",
			m_FolderUserFieldsAnsiCount.getData());

		addBlankLine();

		// Add child nodes to m_FolderUserFieldsUnicodeCount before adding it to our output
		if (m_FolderUserFieldsUnicodeCount && !m_FieldDefinitionsW.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsW)
			{
				auto fieldBlock = block{};
				fieldBlock.setText(L"Field %1!d!\r\n", i++);

				auto szFieldType = flags::InterpretFlags(flagFolderType, fieldDefinition.FieldType);
				fieldBlock.addBlock(
					fieldDefinition.FieldType,
					L"FieldType = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition.FieldType.getData(),
					szFieldType.c_str());
				fieldBlock.addBlock(
					fieldDefinition.FieldNameLength,
					L"FieldNameLength = 0x%1!08X! = %1!d!\r\n",
					fieldDefinition.FieldNameLength.getData());
				fieldBlock.addBlock(
					fieldDefinition.FieldName, L"FieldName = %1!ws!\r\n", fieldDefinition.FieldName.c_str());

				auto szGUID = guid::GUIDToString(fieldDefinition.Common.PropSetGuid);
				fieldBlock.addBlock(fieldDefinition.Common.PropSetGuid, L"PropSetGuid = %1!ws!\r\n", szGUID.c_str());
				auto szFieldcap = flags::InterpretFlags(flagFieldCap, fieldDefinition.Common.fcapm);
				fieldBlock.addBlock(
					fieldDefinition.Common.fcapm,
					L"fcapm = 0x%1!08X! = %2!ws!\r\n",
					fieldDefinition.Common.fcapm.getData(),
					szFieldcap.c_str());
				fieldBlock.addBlock(
					fieldDefinition.Common.dwString,
					L"dwString = 0x%1!08X!\r\n",
					fieldDefinition.Common.dwString.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.dwBitmap,
					L"dwBitmap = 0x%1!08X!\r\n",
					fieldDefinition.Common.dwBitmap.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.dwDisplay,
					L"dwDisplay = 0x%1!08X!\r\n",
					fieldDefinition.Common.dwDisplay.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.iFmt, L"iFmt = 0x%1!08X!\r\n", fieldDefinition.Common.iFmt.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.wszFormulaLength,
					L"wszFormulaLength = 0x%1!04X! = %1!d!\r\n",
					fieldDefinition.Common.wszFormulaLength.getData());
				fieldBlock.addBlock(
					fieldDefinition.Common.wszFormula,
					L"wszFormula = %1!ws!",
					fieldDefinition.Common.wszFormula.c_str());

				m_FolderUserFieldsUnicodeCount.terminateBlock();
				m_FolderUserFieldsUnicodeCount.addBlankLine();
				m_FolderUserFieldsUnicodeCount.addBlock(fieldBlock);
			}
		}

		m_FolderUserFieldsUnicodeCount.terminateBlock();
		addBlock(
			m_FolderUserFieldsUnicodeCount,
			L"FolderUserFieldUnicode.FieldDefinitionCount = %1!d!\r\n",
			m_FolderUserFieldsUnicodeCount.getData());
	}
} // namespace smartview