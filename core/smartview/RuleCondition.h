#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>
#include <core/interpret/guid.h>

namespace smartview
{
	// [MS-OXORULE] 2.2.1.3.1.9 PidTagRuleCondition Property
	// https://msdn.microsoft.com/en-us/library/ee204420(v=exchg.80).aspx

	// [MS-OXORULE] 2.2.4.1.10 PidTagExtendedRuleMessageCondition Property
	// https://msdn.microsoft.com/en-us/library/ee200930(v=exchg.80).aspx

	// RuleRestriction
	// https://msdn.microsoft.com/en-us/library/ee201126(v=exchg.80).aspx

	// [MS-OXCDATA] 2.6.1 NamedProperties Structure
	// http://msdn.microsoft.com/en-us/library/ee158295.aspx
	//   This structure specifies a Property Name
	//
	struct PropertyName : public block
	{
	public:
		PropertyName() = default;
		PropertyName(const PropertyName&) = delete;
		PropertyName& operator=(const PropertyName&) = delete;

		std::shared_ptr<blockT<BYTE>> Kind = emptyT<BYTE>();
		std::shared_ptr<blockT<GUID>> Guid = emptyT<GUID>();
		std::shared_ptr<blockT<DWORD>> LID = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> NameSize = emptyT<BYTE>();
		std::shared_ptr<blockStringW> Name = emptySW();

	private:
		void parse() override;
		void parseBlocks() override{};
	};

	// [MS-OXORULE] 2.2.4.2 NamedPropertyInformation Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/487d4ed5-9896-4d72-a512-e612a9d8147f
	// =====================
	//   This structure specifies named property information for a rule condition
	//
	class NamedPropertyInformation : public block
	{
	private:
		std::shared_ptr<blockT<WORD>> NoOfNamedProps = emptyT<WORD>();
		std::vector<std::shared_ptr<blockT<WORD>>> PropIds;
		std::shared_ptr<blockT<DWORD>> NamedPropertiesSize = emptyT<DWORD>();
		std::vector<std::shared_ptr<PropertyName>> NamedProperties;
		void parse() override
		{
			NoOfNamedProps = blockT<WORD>::parse(parser);
			if (*NoOfNamedProps && *NoOfNamedProps < _MaxEntriesLarge)
			{
				PropIds.reserve(*NoOfNamedProps);
				for (auto i = 0; i < *NoOfNamedProps; i++)
				{
					PropIds.push_back(blockT<WORD>::parse(parser));
				}

				NamedPropertiesSize = blockT<DWORD>::parse(parser);

				NamedProperties.reserve(*NoOfNamedProps);
				for (auto i = 0; i < *NoOfNamedProps; i++)
				{
					NamedProperties.emplace_back(block::parse<PropertyName>(parser, false));
				}
			}
		}
		void parseBlocks() override
		{
			setText(L"NamedPropertyInformation\r\n");
			addChild(NoOfNamedProps, L"Number of named props = 0x%1!04X!\r\n", NoOfNamedProps->getData());
			if (!PropIds.empty())
			{
				terminateBlock();
				addChild(NamedPropertiesSize, L"Named prop size = 0x%1!08X!", NamedPropertiesSize->getData());

				for (size_t i = 0; i < PropIds.size(); i++)
				{
					terminateBlock();
					addChild(NamedProperties[i]);
					NamedProperties[i]->setText(L"Named Prop 0x%1!04X!\r\n", i);

					NamedProperties[i]->addChild(PropIds[i], L"\tPropID = 0x%1!04X!\r\n", PropIds[i]->getData());

					NamedProperties[i]->addChild(
						NamedProperties[i]->Kind, L"\tKind = 0x%1!02X!\r\n", NamedProperties[i]->Kind->getData());
					NamedProperties[i]->addChild(
						NamedProperties[i]->Guid,
						L"\tGuid = %1!ws!\r\n",
						guid::GUIDToString(*NamedProperties[i]->Guid).c_str());

					if (*NamedProperties[i]->Kind == MNID_ID)
					{
						NamedProperties[i]->addChild(
							NamedProperties[i]->LID, L"\tLID = 0x%1!08X!", NamedProperties[i]->LID->getData());
					}
					else if (*NamedProperties[i]->Kind == MNID_STRING)
					{
						NamedProperties[i]->addChild(
							NamedProperties[i]->NameSize,
							L"\tNameSize = 0x%1!02X!\r\n",
							NamedProperties[i]->NameSize->getData());
						NamedProperties[i]->addChild(
							NamedProperties[i]->Name, L"\tName = %1!ws!", NamedProperties[i]->Name->c_str());
					}
				}
			}
		}
	};

	class RuleCondition : public block
	{
	public:
		void Init(bool bExtended) noexcept;

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<NamedPropertyInformation> m_NamedPropertyInformation;
		std::shared_ptr<RestrictionStruct> m_lpRes;
		bool m_bExtended{};
	};
} // namespace smartview