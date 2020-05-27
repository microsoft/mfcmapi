#pragma once

namespace file
{
	enum class exportType
	{
		text = 0,
		msgAnsi = 1,
		msgUnicode = 2,
		eml = 3,
		emlIConverter = 4,
		tnef = 5
	};

	class exporter
	{
	public:
		bool init(CWnd* pParentWnd, bool bMultiSelect, HWND _hWnd, LPADRBOOK _lpAddrBook);
		HRESULT exportMessage(LPMESSAGE lpMessage);

	private:
		exportType exportType{};
		std::wstring szExt{};
		std::wstring szDotExt{};
		std::wstring szFilter{};
		std::wstring dir{};
		HWND hWnd{};
		LPADRBOOK lpAddrBook{};
		bool bPrompt{};
	};
} // namespace file