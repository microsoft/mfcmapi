#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/ViewPane/SplitterPane.h>

namespace viewpane
{
	// Return a pane with a matching paneID.
	template <class T> std::shared_ptr<ViewPane> GetPaneByID(const std::shared_ptr<T> pane, const int id)
	{
		auto splitter = std::dynamic_pointer_cast<SplitterPane>(pane);
		if (splitter)
		{
			if (splitter->GetID() == id) return pane;
			if (splitter->GetPaneOne())
			{
				const auto subpane = GetPaneByID(splitter->GetPaneOne(), id);
				if (subpane) return subpane;
			}

			if (splitter->GetPaneTwo())
			{
				return GetPaneByID(splitter->GetPaneTwo(), id);
			}
		}

		return pane->GetID() == id ? pane : nullptr;
	}

	// Return a pane with a matching nID.
	template <class T> std::shared_ptr<ViewPane> GetPaneByNID(const std::shared_ptr<T> pane, const UINT id)
	{
		auto splitter = std::dynamic_pointer_cast<SplitterPane>(pane);
		if (splitter)
		{
			if (splitter->GetNID() == id) return pane;
			if (splitter->GetPaneOne())
			{
				const auto subpane = GetPaneByNID(splitter->GetPaneOne(), id);
				if (subpane) return subpane;
			}

			if (splitter->GetPaneTwo())
			{
				return GetPaneByNID(splitter->GetPaneTwo(), id);
			}
		}

		return pane->GetNID() == id ? pane : nullptr;
	}
} // namespace viewpane