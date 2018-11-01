#pragma once
#include <Interpret/SmartView/block/block.h>

namespace smartview
{
	template <class T> class blockT : public block
	{
	public:
		blockT() { memset(&data, 0, sizeof(data)); }
		void setData(const T& _data) { data = _data; }
		T getData() const { return data; }
		operator T&() { return data; }
		operator T() const { return data; }
		T& operator=(const blockT<T>& _data)
		{
			data = _data.data;
			cb = _data.getSize();
			offset = _data.getOffset();
			return *this;
		}
		T& operator=(const T& _data)
		{
			data = _data;
			return *this;
		}

	private:
		T data;
	};
} // namespace smartview