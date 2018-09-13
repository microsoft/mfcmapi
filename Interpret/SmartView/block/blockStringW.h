#pragma once
#include <Interpret/SmartView/block/block.h>

namespace smartview
{
	class blockStringW : public block
	{
	public:
		void setData(const std::wstring& _data) { data = _data; }
		operator std::wstring&() { return data; }
		_Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		size_t length() const noexcept { return data.length(); }
		bool empty() const noexcept { return data.empty(); }

	private:
		std::wstring data;
	};
} // namespace smartview