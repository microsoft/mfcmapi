#pragma once
#include <Interpret/SmartView/block/block.h>

namespace smartview
{
	class blockBytes : public block
	{
	public:
		void setData(const std::vector<BYTE>& data) { _data = data; }
		std::vector<BYTE> getData() const { return _data; }
		operator const std::vector<BYTE>&() const { return _data; }
		operator const std::vector<BYTE>() const { return _data; }
		size_t size() const noexcept { return _data.size(); }
		bool empty() const noexcept { return _data.empty(); }
		const BYTE* data() const noexcept { return _data.data(); }
		BYTE operator[](const size_t _Pos) { return _data[_Pos]; }

	private:
		std::vector<BYTE> _data;
	};
} // namespace smartview