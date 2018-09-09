#pragma once
#include <Interpret/SmartView/block/block.h>

namespace smartview
{
	class blockStringA : public block
	{
	public:
		void setData(const std::string& _data) { data = _data; }
		std::string getData() const { return data; }
		operator std::string&() { return data; }
		operator std::string() const { return data; }
		_Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		size_t length() const noexcept { return data.length(); }
		bool empty() const noexcept { return data.empty(); }

	private:
		std::string data;
	};
} // namespace smartview