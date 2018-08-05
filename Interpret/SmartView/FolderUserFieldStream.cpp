#include <StdAfx.h>
#include <Interpret/SmartView/FolderUserFieldStream.h>
#include <Interpret/Guids.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	FolderUserFieldStream::FolderUserFieldStream() {}

	void FolderUserFieldStream::Parse()
	{
		m_FolderUserFieldsAnsiCount = m_Parser.GetBlock<DWORD>();
		if (m_FolderUserFieldsAnsiCount.getData() && m_FolderUserFieldsAnsiCount.getData() < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_FolderUserFieldsAnsiCount.getData(); i++)
			{
				FolderFieldDefinitionA folderFieldDefinitionA;
				folderFieldDefinitionA.FieldType = m_Parser.GetBlock<DWORD>();
				folderFieldDefinitionA.FieldNameLength = m_Parser.GetBlock<WORD>();

				if (folderFieldDefinitionA.FieldNameLength.getData() &&
					folderFieldDefinitionA.FieldNameLength.getData() < _MaxEntriesSmall)
				{
					folderFieldDefinitionA.FieldName =
						m_Parser.GetBlockStringA(folderFieldDefinitionA.FieldNameLength.getData());
				}

				folderFieldDefinitionA.Common = BinToFolderFieldDefinitionCommon();
				m_FieldDefinitionsA.push_back(folderFieldDefinitionA);
			}
		}

		m_FolderUserFieldsUnicodeCount = m_Parser.GetBlock<DWORD>();
		if (m_FolderUserFieldsUnicodeCount.getData() && m_FolderUserFieldsUnicodeCount.getData() < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_FolderUserFieldsUnicodeCount.getData(); i++)
			{
				FolderFieldDefinitionW folderFieldDefinitionW;
				folderFieldDefinitionW.FieldType = m_Parser.GetBlock<DWORD>();
				folderFieldDefinitionW.FieldNameLength = m_Parser.GetBlock<WORD>();

				if (folderFieldDefinitionW.FieldNameLength.getData() &&
					folderFieldDefinitionW.FieldNameLength.getData() < _MaxEntriesSmall)
				{
					folderFieldDefinitionW.FieldName =
						m_Parser.GetBlockStringW(folderFieldDefinitionW.FieldNameLength.getData());
				}

				folderFieldDefinitionW.Common = BinToFolderFieldDefinitionCommon();
				m_FieldDefinitionsW.push_back(folderFieldDefinitionW);
			}
		}

		ParseBlocks();
	}

	FolderFieldDefinitionCommon FolderUserFieldStream::BinToFolderFieldDefinitionCommon()
	{
		FolderFieldDefinitionCommon common;

		common.PropSetGuid = m_Parser.GetBlock<GUID>();
		common.fcapm = m_Parser.GetBlock<DWORD>();
		common.dwString = m_Parser.GetBlock<DWORD>();
		common.dwBitmap = m_Parser.GetBlock<DWORD>();
		common.dwDisplay = m_Parser.GetBlock<DWORD>();
		common.iFmt = m_Parser.GetBlock<DWORD>();
		common.wszFormulaLength = m_Parser.GetBlock<WORD>();
		if (common.wszFormulaLength.getData() && common.wszFormulaLength.getData() < _MaxEntriesLarge)
		{
			common.wszFormula = m_Parser.GetBlockStringW(common.wszFormulaLength.getData());
		}

		return common;
	}

	void FolderUserFieldStream::ParseBlocks()
	{
		addHeader(L"Folder User Field Stream\r\n");

		// Add child nodes to m_FolderUserFieldsAnsiCount before adding it to our output
		if (m_FolderUserFieldsAnsiCount.getData() && !m_FieldDefinitionsA.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsA)
			{
				m_FolderUserFieldsAnsiCount.addHeader(L"\r\n\r\n");
				m_FolderUserFieldsAnsiCount.addHeader(strings::formatmessage(L"Field %1!d!\r\n", i++));

				auto szFieldType = interpretprop::InterpretFlags(flagFolderType, fieldDefinition.FieldType.getData());
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.FieldType,
					strings::formatmessage(
						L"FieldType = 0x%1!08X! = %2!ws!\r\n",
						fieldDefinition.FieldType.getData(),
						szFieldType.c_str()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.FieldNameLength,
					strings::formatmessage(
						L"FieldNameLength = 0x%1!08X! = %1!d!\r\n", fieldDefinition.FieldNameLength.getData()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.FieldName,
					strings::formatmessage(L"FieldName = %1!hs!\r\n", fieldDefinition.FieldName.getData().c_str()));

				auto szGUID = guid::GUIDToString(fieldDefinition.Common.PropSetGuid.getData());
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.PropSetGuid,
					strings::formatmessage(L"PropSetGuid = %1!ws!\r\n", szGUID.c_str()));
				auto szFieldcap = interpretprop::InterpretFlags(flagFieldCap, fieldDefinition.Common.fcapm.getData());
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.fcapm,
					strings::formatmessage(
						L"fcapm = 0x%1!08X! = %2!ws!\r\n", fieldDefinition.Common.fcapm.getData(), szFieldcap.c_str()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.dwString,
					strings::formatmessage(L"dwString = 0x%1!08X!\r\n", fieldDefinition.Common.dwString.getData()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.dwBitmap,
					strings::formatmessage(L"dwBitmap = 0x%1!08X!\r\n", fieldDefinition.Common.dwBitmap.getData()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.dwDisplay,
					strings::formatmessage(L"dwDisplay = 0x%1!08X!\r\n", fieldDefinition.Common.dwDisplay.getData()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.iFmt,
					strings::formatmessage(L"iFmt = 0x%1!08X!\r\n", fieldDefinition.Common.iFmt.getData()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.wszFormulaLength,
					strings::formatmessage(
						L"wszFormulaLength = 0x%1!04X! = %1!d!\r\n",
						fieldDefinition.Common.wszFormulaLength.getData()));
				m_FolderUserFieldsAnsiCount.addBlock(
					fieldDefinition.Common.wszFormula,
					strings::formatmessage(
						L"wszFormula = %1!ws!", fieldDefinition.Common.wszFormula.getData().c_str()));
			}
		}

		addBlock(
			m_FolderUserFieldsAnsiCount,
			strings::formatmessage(
				L"FolderUserFieldAnsi.FieldDefinitionCount = %1!d!", m_FolderUserFieldsAnsiCount.getData()));

		addHeader(L"\r\n\r\n");

		// Add child nodes to m_FolderUserFieldsUnicodeCount before adding it to our output
		if (m_FolderUserFieldsUnicodeCount.getData() && !m_FieldDefinitionsW.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsW)
			{
				m_FolderUserFieldsUnicodeCount.addHeader(L"\r\n\r\n");
				m_FolderUserFieldsUnicodeCount.addHeader(strings::formatmessage(L"Field %1!d!\r\n", i++));

				auto szFieldType = interpretprop::InterpretFlags(flagFolderType, fieldDefinition.FieldType.getData());
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.FieldType,
					strings::formatmessage(
						L"FieldType = 0x%1!08X! = %2!ws!\r\n",
						fieldDefinition.FieldType.getData(),
						szFieldType.c_str()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.FieldNameLength,
					strings::formatmessage(
						L"FieldNameLength = 0x%1!08X! = %1!d!\r\n", fieldDefinition.FieldNameLength.getData()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.FieldName,
					strings::formatmessage(L"FieldName = %1!ws!\r\n", fieldDefinition.FieldName.getData().c_str()));

				auto szGUID = guid::GUIDToString(fieldDefinition.Common.PropSetGuid.getData());
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.PropSetGuid,
					strings::formatmessage(L"PropSetGuid = %1!ws!\r\n", szGUID.c_str()));
				auto szFieldcap = interpretprop::InterpretFlags(flagFieldCap, fieldDefinition.Common.fcapm.getData());
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.fcapm,
					strings::formatmessage(
						L"fcapm = 0x%1!08X! = %2!ws!\r\n", fieldDefinition.Common.fcapm.getData(), szFieldcap.c_str()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.dwString,
					strings::formatmessage(L"dwString = 0x%1!08X!\r\n", fieldDefinition.Common.dwString.getData()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.dwBitmap,
					strings::formatmessage(L"dwBitmap = 0x%1!08X!\r\n", fieldDefinition.Common.dwBitmap.getData()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.dwDisplay,
					strings::formatmessage(L"dwDisplay = 0x%1!08X!\r\n", fieldDefinition.Common.dwDisplay.getData()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.iFmt,
					strings::formatmessage(L"iFmt = 0x%1!08X!\r\n", fieldDefinition.Common.iFmt.getData()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.wszFormulaLength,
					strings::formatmessage(
						L"wszFormulaLength = 0x%1!04X! = %1!d!\r\n",
						fieldDefinition.Common.wszFormulaLength.getData()));
				m_FolderUserFieldsUnicodeCount.addBlock(
					fieldDefinition.Common.wszFormula,
					strings::formatmessage(
						L"wszFormula = %1!ws!", fieldDefinition.Common.wszFormula.getData().c_str()));
			}
		}

		addBlock(
			m_FolderUserFieldsUnicodeCount,
			strings::formatmessage(
				L"FolderUserFieldUnicode.FieldDefinitionCount = %1!d!", m_FolderUserFieldsUnicodeCount.getData()));
	}
}