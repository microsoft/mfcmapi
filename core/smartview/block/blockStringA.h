#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class blockStringA : public block
	{
	public:
		void setData(const std::string& _data) { data = _data; }
		operator const std::string&() const { return data; }
		_NODISCARD _Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::string::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

	private:
		std::string data;
	};
} // namespace smartview