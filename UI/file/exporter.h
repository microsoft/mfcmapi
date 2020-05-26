
namespace exporter
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

	struct exportOptions
	{
		std::wstring szExt{};
		std::wstring szDotExt{};
		std::wstring szFilter{};
	};

	exportOptions getExportOptions(exporter::exportType exportType);
	HRESULT exportMessage(
		exporter::exportType exportType,
		std::wstring filename,
		LPMESSAGE lpMessage,
		HWND hWnd,
		LPADRBOOK lpAddrBook);
} // namespace exporter