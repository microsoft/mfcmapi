#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class blockStringW : public block
	{
	public:
		void setData(const std::wstring& _data) { data = _data; }
		operator const std::wstring&() const { return data; }
		_NODISCARD _Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::wstring::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

	private:
		std::wstring data;
	};
} // namespace smartview