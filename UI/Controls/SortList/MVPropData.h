#pragma once
#include <MAPIDefS.h>
#include <UI/Controls/SortList/Data.h>

namespace controls
{
	namespace sortlistdata
	{
		class MVPropData : public IData
		{
		public:
			MVPropData(_In_opt_ const _SPropValue* lpProp, ULONG iProp);
			MVPropData(_In_opt_ const _SPropValue* lpProp);
			_PV m_val{};

		private:
			std::string m_lpszA;
			std::wstring m_lpszW;
			std::vector<BYTE> m_lpBin;
		};
	} // namespace sortlistdata
} // namespace controls