#include <StdAfx.h>
#include <UI/Dialogs/Editors/RestrictEditor.h>
#include <UI/Dialogs/Editors/PropertyEditor.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/extraPropTags.h>
#include <core/sortlistdata/resData.h>
#include <core/sortlistdata/commentData.h>
#include <core/sortlistdata/binaryData.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog
{
	namespace editor
	{
		static std::wstring COMPCLASS = L"CResCompareEditor"; // STRING_OK
		class CResCompareEditor : public CEditor
		{
		public:
			CResCompareEditor(_In_ CWnd* pParentWnd, ULONG ulRelop, ULONG ulPropTag1, ULONG ulPropTag2);

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;
		};

		CResCompareEditor::CResCompareEditor(_In_ CWnd* pParentWnd, ULONG ulRelop, ULONG ulPropTag1, ULONG ulPropTag2)
			: CEditor(pParentWnd, IDS_RESED, IDS_RESEDCOMPPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(COMPCLASS);

			SetPromptPostFix(flags::AllFlagsToString(flagRelop, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELOP, false));
			SetHex(0, ulRelop);
			const auto szFlags = flags::InterpretFlags(flagRelop, ulRelop);
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELOP, szFlags, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG1, false));
			SetHex(2, ulPropTag1);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				3, IDS_ULPROPTAG1, proptags::TagToString(ulPropTag1, nullptr, false, true), true));

			AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_ULPROPTAG2, false));
			SetHex(4, ulPropTag2);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				5, IDS_ULPROPTAG1, proptags::TagToString(ulPropTag2, nullptr, false, true), true));
		}

		_Check_return_ ULONG CResCompareEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == 0)
			{
				SetStringW(1, flags::InterpretFlags(flagRelop, GetHex(0)));
			}
			else if (paneID == 2)
			{
				SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
			}
			else if (paneID == 4)
			{
				SetStringW(5, proptags::TagToString(GetPropTag(4), nullptr, false, true));
			}

			return paneID;
		}

		// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
		// So all memory detached from this class is owned by a parent and must not be freed manually
		static std::wstring CONTENTCLASS = L"CResCombinedEditor"; // STRING_OK
		class CResCombinedEditor : public CEditor
		{
		public:
			CResCombinedEditor(
				_In_ CWnd* pParentWnd,
				ULONG ulResType,
				ULONG ulCompare,
				ULONG ulPropTag,
				_In_ const _SPropValue* lpProp,
				_In_ LPVOID lpAllocParent);

			void OnEditAction1() override;
			_Check_return_ LPSPropValue DetachModifiedSPropValue();

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;

			ULONG m_ulResType;
			LPVOID m_lpAllocParent;
			const _SPropValue* m_lpOldProp;
			LPSPropValue m_lpNewProp;
		};

		CResCombinedEditor::CResCombinedEditor(
			_In_ CWnd* pParentWnd,
			ULONG ulResType,
			ULONG ulCompare,
			ULONG ulPropTag,
			_In_ const _SPropValue* lpProp,
			_In_ LPVOID lpAllocParent)
			: CEditor(
				  pParentWnd,
				  IDS_RESED,
				  ulResType == RES_CONTENT ? IDS_RESEDCONTPROMPT : // Content case
					  ulResType == RES_PROPERTY ? IDS_RESEDPROPPROMPT : // Property case
						  0, // else case
				  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
				  IDS_ACTIONEDITPROP,
				  NULL,
				  NULL)
		{
			TRACE_CONSTRUCTOR(CONTENTCLASS);

			m_ulResType = ulResType;
			m_lpOldProp = lpProp;
			m_lpNewProp = nullptr;
			m_lpAllocParent = lpAllocParent;

			std::wstring szFlags;

			if (RES_CONTENT == m_ulResType)
			{
				SetPromptPostFix(flags::AllFlagsToString(flagFuzzyLevel, true));

				AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULFUZZYLEVEL, false));
				SetHex(0, ulCompare);
				szFlags = flags::InterpretFlags(flagFuzzyLevel, ulCompare);
				AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_ULFUZZYLEVEL, szFlags, true));
			}
			else if (RES_PROPERTY == m_ulResType)
			{
				SetPromptPostFix(flags::AllFlagsToString(flagRelop, false));
				AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELOP, false));
				SetHex(0, ulCompare);
				szFlags = flags::InterpretFlags(flagRelop, ulCompare);
				AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELOP, szFlags, true));
			}

			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG, false));
			SetHex(2, ulPropTag);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				3, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));

			AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_LPPROPULPROPTAG, false));
			if (lpProp) SetHex(4, lpProp->ulPropTag);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				5,
				IDS_LPPROPULPROPTAG,
				lpProp ? proptags::TagToString(lpProp->ulPropTag, nullptr, false, true) : strings::emptystring,
				true));

			std::wstring szProp;
			std::wstring szAltProp;
			if (lpProp) property::parseProperty(lpProp, &szProp, &szAltProp);
			AddPane(viewpane::TextPane::CreateMultiLinePane(6, IDS_LPPROP, szProp, true));
			AddPane(viewpane::TextPane::CreateMultiLinePane(7, IDS_LPPROPALTVIEW, szAltProp, true));
		}

		_Check_return_ ULONG CResCombinedEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			std::wstring szFlags;
			if (paneID == 0)
			{
				if (RES_CONTENT == m_ulResType)
				{
					SetStringW(1, flags::InterpretFlags(flagFuzzyLevel, GetHex(0)));
				}
				else if (RES_PROPERTY == m_ulResType)
				{
					SetStringW(1, flags::InterpretFlags(flagRelop, GetHex(0)));
				}
			}
			else if (paneID == 2)
			{
				SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
			}
			else if (paneID == 4)
			{
				SetStringW(5, proptags::TagToString(GetPropTag(4), nullptr, false, true));
				m_lpOldProp = nullptr;
				m_lpNewProp = nullptr;
				SetStringW(6, L"");
				SetStringW(7, L"");
			}

			return paneID;
		}

		_Check_return_ LPSPropValue CResCombinedEditor::DetachModifiedSPropValue()
		{
			const auto lpRet = m_lpNewProp;
			m_lpNewProp = nullptr;
			return lpRet;
		}

		void CResCombinedEditor::OnEditAction1()
		{
			if (!m_lpAllocParent) return;

			auto lpEditProp = m_lpOldProp;
			LPSPropValue lpOutProp = nullptr;
			if (m_lpNewProp) lpEditProp = m_lpNewProp;

			const auto hRes = WC_H(DisplayPropertyEditor(
				this,
				IDS_PROPEDITOR,
				NULL,
				false,
				m_lpAllocParent,
				NULL,
				GetPropTag(4),
				false,
				lpEditProp,
				&lpOutProp));

			// Since m_lpNewProp was owned by an m_lpAllocParent, we don't free it directly
			if (hRes == S_OK && lpOutProp)
			{
				m_lpNewProp = lpOutProp;
				std::wstring szProp;
				std::wstring szAltProp;

				property::parseProperty(m_lpNewProp, &szProp, &szAltProp);
				SetStringW(6, szProp);
				SetStringW(7, szAltProp);
			}
		}

		static std::wstring BITMASKCLASS = L"CResBitmaskEditor"; // STRING_OK
		class CResBitmaskEditor : public CEditor
		{
		public:
			CResBitmaskEditor(_In_ CWnd* pParentWnd, ULONG relBMR, ULONG ulPropTag, ULONG ulMask);

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;
		};

		CResBitmaskEditor::CResBitmaskEditor(_In_ CWnd* pParentWnd, ULONG relBMR, ULONG ulPropTag, ULONG ulMask)
			: CEditor(pParentWnd, IDS_RESED, IDS_RESEDBITPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(BITMASKCLASS);

			SetPromptPostFix(flags::AllFlagsToString(flagBitmask, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELBMR, false));
			SetHex(0, relBMR);
			const auto szFlags = flags::InterpretFlags(flagBitmask, relBMR);
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELBMR, szFlags, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG, false));
			SetHex(2, ulPropTag);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				3, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));

			AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_MASK, false));
			SetHex(4, ulMask);
		}

		_Check_return_ ULONG CResBitmaskEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == 0)
			{
				SetStringW(1, flags::InterpretFlags(flagBitmask, GetHex(0)));
			}
			else if (paneID == 2)
			{
				SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
			}

			return paneID;
		}

		static std::wstring SIZECLASS = L"CResSizeEditor"; // STRING_OK
		class CResSizeEditor : public CEditor
		{
		public:
			CResSizeEditor(_In_ CWnd* pParentWnd, ULONG relop, ULONG ulPropTag, ULONG cb);

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;
		};

		CResSizeEditor::CResSizeEditor(_In_ CWnd* pParentWnd, ULONG relop, ULONG ulPropTag, ULONG cb)
			: CEditor(pParentWnd, IDS_RESED, IDS_RESEDSIZEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(SIZECLASS);

			SetPromptPostFix(flags::AllFlagsToString(flagRelop, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELOP, false));
			SetHex(0, relop);
			const auto szFlags = flags::InterpretFlags(flagRelop, relop);
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELOP, szFlags, true));

			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG, false));
			SetHex(2, ulPropTag);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				3, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));

			AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_CB, false));
			SetHex(4, cb);
		}

		_Check_return_ ULONG CResSizeEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == 0)
			{
				SetStringW(1, flags::InterpretFlags(flagRelop, GetHex(0)));
			}
			else if (paneID == 2)
			{
				SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
			}

			return paneID;
		}

		static std::wstring EXISTCLASS = L"CResExistEditor"; // STRING_OK
		class CResExistEditor : public CEditor
		{
		public:
			CResExistEditor(_In_ CWnd* pParentWnd, ULONG ulPropTag);

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;
		};

		CResExistEditor::CResExistEditor(_In_ CWnd* pParentWnd, ULONG ulPropTag)
			: CEditor(pParentWnd, IDS_RESED, IDS_RESEDEXISTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(EXISTCLASS);

			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULPROPTAG, false));
			SetHex(0, ulPropTag);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				1, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));
		}

		_Check_return_ ULONG CResExistEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == 0)
			{
				SetStringW(1, proptags::TagToString(GetPropTag(0), nullptr, false, true));
			}

			return paneID;
		}

		// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
		// So all memory detached from this class is owned by a parent and must not be freed manually
		static std::wstring SUBRESCLASS = L"CResSubResEditor"; // STRING_OK
		class CResSubResEditor : public CEditor
		{
		public:
			CResSubResEditor(
				_In_ CWnd* pParentWnd,
				ULONG ulSubObject,
				_In_ const _SRestriction* lpRes,
				_In_ LPVOID lpAllocParent);

			void OnEditAction1() override;
			_Check_return_ LPSRestriction DetachModifiedSRestriction();

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;

			LPVOID m_lpAllocParent;
			const _SRestriction* m_lpOldRes;
			LPSRestriction m_lpNewRes;
		};

		CResSubResEditor::CResSubResEditor(
			_In_ CWnd* pParentWnd,
			ULONG ulSubObject,
			_In_ const _SRestriction* lpRes,
			_In_ LPVOID lpAllocParent)
			: CEditor(
				  pParentWnd,
				  IDS_SUBRESED,
				  IDS_RESEDSUBPROMPT,
				  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
				  IDS_ACTIONEDITRES,
				  NULL,
				  NULL)
		{
			TRACE_CONSTRUCTOR(SUBRESCLASS);

			m_lpOldRes = lpRes;
			m_lpNewRes = nullptr;
			m_lpAllocParent = lpAllocParent;

			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULSUBOBJECT, false));
			SetHex(0, ulSubObject);
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				1, IDS_ULSUBOBJECT, proptags::TagToString(ulSubObject, nullptr, false, true), true));

			AddPane(viewpane::TextPane::CreateMultiLinePane(
				2, IDS_LPRES, property::RestrictionToString(lpRes, nullptr), true));
		}

		_Check_return_ ULONG CResSubResEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == 0)
			{
				SetStringW(1, proptags::TagToString(GetPropTag(0), nullptr, false, true));
			}

			return paneID;
		}

		_Check_return_ LPSRestriction CResSubResEditor::DetachModifiedSRestriction()
		{
			const auto lpRet = m_lpNewRes;
			m_lpNewRes = nullptr;
			return lpRet;
		}

		void CResSubResEditor::OnEditAction1()
		{
			CRestrictEditor ResEdit(this, m_lpAllocParent, m_lpNewRes ? m_lpNewRes : m_lpOldRes);

			if (!ResEdit.DisplayDialog()) return;

			// Since m_lpNewRes was owned by an m_lpAllocParent, we don't free it directly
			m_lpNewRes = ResEdit.DetachModifiedSRestriction();

			SetStringW(2, property::RestrictionToString(m_lpNewRes, nullptr));
		}

		// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
		// So all memory detached from this class is owned by a parent and must not be freed manually
		static std::wstring ANDORCLASS = L"CResAndOrEditor"; // STRING_OK
		class CResAndOrEditor : public CEditor
		{
		public:
			CResAndOrEditor(_In_ CWnd* pParentWnd, _In_ const _SRestriction* lpRes, _In_ LPVOID lpAllocParent);

			_Check_return_ LPSRestriction DetachModifiedSRestrictionArray();
			_Check_return_ ULONG GetResCount() const;
			_Check_return_ bool
			DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData) override;

		private:
			BOOL OnInitDialog() override;
			void InitListFromRestriction(ULONG ulListNum, _In_ const _SRestriction* lpRes) const;
			void OnOK() override;

			LPVOID m_lpAllocParent;
			const _SRestriction* m_lpRes;
			LPSRestriction m_lpNewResArray;
			ULONG m_ulNewResCount;
		};

		CResAndOrEditor::CResAndOrEditor(
			_In_ CWnd* pParentWnd,
			_In_ const _SRestriction* lpRes,
			_In_ LPVOID lpAllocParent)
			: CEditor(pParentWnd, IDS_RESED, IDS_RESEDANDORPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(ANDORCLASS);
			m_lpRes = lpRes;
			m_lpNewResArray = nullptr;
			m_ulNewResCount = NULL;
			m_lpAllocParent = lpAllocParent;

			AddPane(viewpane::ListPane::Create(0, IDS_SUBRESTRICTIONS, false, false, ListEditCallBack(this)));
			SetListID(0);
		}

		// Used to call functions which need to be called AFTER controls are created
		BOOL CResAndOrEditor::OnInitDialog()
		{
			const auto bRet = CEditor::OnInitDialog();

			InitListFromRestriction(0, m_lpRes);

			UpdateButtons();

			return bRet;
		}

		_Check_return_ LPSRestriction CResAndOrEditor::DetachModifiedSRestrictionArray()
		{
			const auto lpRet = m_lpNewResArray;
			m_lpNewResArray = nullptr;
			return lpRet;
		}

		_Check_return_ ULONG CResAndOrEditor::GetResCount() const { return m_ulNewResCount; }

		void CResAndOrEditor::InitListFromRestriction(ULONG ulListNum, _In_ const _SRestriction* lpRes) const
		{
			ClearList(ulListNum);
			InsertColumn(ulListNum, 0, IDS_SHARP);

			if (!lpRes) return;
			switch (lpRes->rt)
			{
			case RES_AND:
				InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

				for (ULONG paneID = 0; paneID < lpRes->res.resAnd.cRes; paneID++)
				{
					auto lpData = InsertListRow(ulListNum, paneID, std::to_wstring(paneID));
					if (lpData)
					{
						lpData->InitializeRes(&lpRes->res.resAnd.lpRes[paneID]);
						SetListString(
							ulListNum,
							paneID,
							1,
							property::RestrictionToString(&lpRes->res.resAnd.lpRes[paneID], nullptr));
					}
				}
				break;
			case RES_OR:
				InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

				for (ULONG paneID = 0; paneID < lpRes->res.resOr.cRes; paneID++)
				{
					auto lpData = InsertListRow(ulListNum, paneID, std::to_wstring(paneID));
					if (lpData)
					{
						lpData->InitializeRes(&lpRes->res.resOr.lpRes[paneID]);
						SetListString(
							ulListNum,
							paneID,
							1,
							property::RestrictionToString(&lpRes->res.resOr.lpRes[paneID], nullptr));
					}
				}
				break;
			}
			ResizeList(ulListNum, false);
		}

		_Check_return_ bool
		CResAndOrEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData)
		{
			if (!lpData) return false;
			if (!lpData->Res())
			{
				lpData->InitializeRes(nullptr);
			}

			if (!lpData->Res()) return false;

			const auto lpSourceRes = lpData->Res()->m_lpNewRes ? lpData->Res()->m_lpNewRes : lpData->Res()->m_lpOldRes;

			CRestrictEditor MyResEditor(this, m_lpAllocParent,
										lpSourceRes); // pass source res into editor
			if (!MyResEditor.DisplayDialog()) return false;
			// Since lpData->data.Res.lpNewRes was owned by an m_lpAllocParent, we don't free it directly
			lpData->Res()->m_lpNewRes = MyResEditor.DetachModifiedSRestriction();
			SetListString(ulListNum, iItem, 1, property::RestrictionToString(lpData->Res()->m_lpNewRes, nullptr));
			return true;
		}

		// Create our LPSRestriction array from the dialog here
		void CResAndOrEditor::OnOK()
		{
			CMyDialog::OnOK(); // don't need to call CEditor::OnOK

			const auto ulNewResCount = GetListCount(0);

			if (ulNewResCount > ULONG_MAX / sizeof(SRestriction)) return;
			auto lpNewResArray = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * ulNewResCount, m_lpAllocParent);
			if (lpNewResArray)
			{
				for (ULONG paneID = 0; paneID < ulNewResCount; paneID++)
				{
					const auto lpData = GetListRowData(0, paneID);
					if (lpData && lpData->Res())
					{
						if (lpData->Res()->m_lpNewRes)
						{
							memcpy(&lpNewResArray[paneID], lpData->Res()->m_lpNewRes, sizeof(SRestriction));
							memset(lpData->Res()->m_lpNewRes, 0, sizeof(SRestriction));
						}
						else
						{
							EC_H_S(mapi::HrCopyRestrictionArray(
								lpData->Res()->m_lpOldRes, m_lpAllocParent, 1, &lpNewResArray[paneID]));
						}
					}
				}
				m_ulNewResCount = ulNewResCount;
				m_lpNewResArray = lpNewResArray;
			}
		}

		// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
		// So all memory detached from this class is owned by a parent and must not be freed manually
		static std::wstring COMMENTCLASS = L"CResCommentEditor"; // STRING_OK
		class CResCommentEditor : public CEditor
		{
		public:
			CResCommentEditor(_In_ CWnd* pParentWnd, _In_ const _SRestriction* lpRes, _In_ LPVOID lpAllocParent);

			_Check_return_ LPSRestriction DetachModifiedSRestriction();
			_Check_return_ LPSPropValue DetachModifiedSPropValue();
			_Check_return_ ULONG GetSPropValueCount() const;
			_Check_return_ bool
			DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData) override;

		private:
			void OnEditAction1() override;
			BOOL OnInitDialog() override;
			void
			InitListFromPropArray(ULONG ulListNum, ULONG cProps, _In_count_(cProps) const _SPropValue* lpProps) const;
			_Check_return_ const _SRestriction* GetSourceRes() const;
			void OnOK() override;

			LPVOID m_lpAllocParent;
			const _SRestriction* m_lpSourceRes;

			LPSRestriction m_lpNewCommentRes;
			LPSPropValue m_lpNewCommentProp;
			ULONG m_ulNewCommentProp{};
		};

		CResCommentEditor::CResCommentEditor(
			_In_ CWnd* pParentWnd,
			_In_ const _SRestriction* lpRes,
			_In_ LPVOID lpAllocParent)
			: CEditor(
				  pParentWnd,
				  IDS_COMMENTRESED,
				  IDS_RESEDCOMMENTPROMPT,
				  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
				  IDS_ACTIONEDITRES,
				  NULL,
				  NULL)
		{
			TRACE_CONSTRUCTOR(COMMENTCLASS);

			m_lpSourceRes = lpRes;
			m_lpNewCommentRes = nullptr;
			m_lpNewCommentProp = nullptr;
			m_lpAllocParent = lpAllocParent;

			AddPane(viewpane::ListPane::CreateCollapsibleListPane(
				0, IDS_SUBRESTRICTION, false, false, ListEditCallBack(this)));
			SetListID(0);
			AddPane(viewpane::TextPane::CreateMultiLinePane(
				1,
				IDS_RESTRICTIONTEXT,
				property::RestrictionToString(m_lpSourceRes->res.resComment.lpRes, nullptr),
				true));
		}

		// Used to call functions which need to be called AFTER controls are created
		BOOL CResCommentEditor::OnInitDialog()
		{
			const auto bRet = CEditor::OnInitDialog();

			InitListFromPropArray(0, m_lpSourceRes->res.resComment.cValues, m_lpSourceRes->res.resComment.lpProp);

			UpdateButtons();

			return bRet;
		}

		_Check_return_ const _SRestriction* CResCommentEditor::GetSourceRes() const
		{
			if (m_lpNewCommentRes) return m_lpNewCommentRes;
			if (m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes) return m_lpSourceRes->res.resComment.lpRes;
			return nullptr;
		}

		_Check_return_ LPSRestriction CResCommentEditor::DetachModifiedSRestriction()
		{
			const auto lpRet = m_lpNewCommentRes;
			m_lpNewCommentRes = nullptr;
			return lpRet;
		}

		_Check_return_ LPSPropValue CResCommentEditor::DetachModifiedSPropValue()
		{
			const auto lpRet = m_lpNewCommentProp;
			m_lpNewCommentProp = nullptr;
			return lpRet;
		}

		_Check_return_ ULONG CResCommentEditor::GetSPropValueCount() const { return m_ulNewCommentProp; }

		void CResCommentEditor::InitListFromPropArray(
			ULONG ulListNum,
			ULONG cProps,
			_In_count_(cProps) const _SPropValue* lpProps) const
		{
			ClearList(ulListNum);

			InsertColumn(ulListNum, 0, IDS_SHARP);
			std::wstring szProp;
			std::wstring szAltProp;
			InsertColumn(ulListNum, 1, IDS_PROPERTY);
			InsertColumn(ulListNum, 2, IDS_VALUE);
			InsertColumn(ulListNum, 3, IDS_ALTERNATEVIEW);

			for (ULONG paneID = 0; paneID < cProps; paneID++)
			{
				auto lpData = InsertListRow(ulListNum, paneID, std::to_wstring(paneID));
				if (lpData)
				{
					lpData->InitializeComment(&lpProps[paneID]);
					SetListString(
						ulListNum, paneID, 1, proptags::TagToString(lpProps[paneID].ulPropTag, nullptr, false, true));
					property::parseProperty(&lpProps[paneID], &szProp, &szAltProp);
					SetListString(ulListNum, paneID, 2, szProp);
					SetListString(ulListNum, paneID, 3, szAltProp);
				}
			}

			ResizeList(ulListNum, false);
		}

		void CResCommentEditor::OnEditAction1()
		{
			const auto lpSourceRes = GetSourceRes();

			CRestrictEditor MyResEditor(this, m_lpAllocParent,
										lpSourceRes); // pass source res into editor
			if (!MyResEditor.DisplayDialog()) return;

			// Since m_lpNewCommentRes was owned by an m_lpAllocParent, we don't free it directly
			m_lpNewCommentRes = MyResEditor.DetachModifiedSRestriction();
			SetStringW(1, property::RestrictionToString(m_lpNewCommentRes, nullptr));
		}

		_Check_return_ bool
		CResCommentEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData)
		{
			if (!m_lpAllocParent) return false;
			if (!lpData) return false;
			if (!lpData->Res())
			{
				lpData->InitializeComment(nullptr);
			}

			if (!lpData->Comment()) return false;

			auto lpSourceProp =
				lpData->Comment()->m_lpNewProp ? lpData->Comment()->m_lpNewProp : lpData->Comment()->m_lpOldProp;

			auto sProp = SPropValue{};

			if (!lpSourceProp)
			{
				CEditor MyTag(this, IDS_TAG, IDS_TAGPROMPT, true);

				MyTag.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_TAG, false));

				if (!MyTag.DisplayDialog()) return false;
				sProp.ulPropTag = MyTag.GetHex(0);
				lpSourceProp = &sProp;
			}

			const auto hRes = WC_H(DisplayPropertyEditor(
				this,
				IDS_PROPEDITOR,
				NULL,
				false,
				m_lpAllocParent,
				NULL,
				NULL,
				false,
				lpSourceProp,
				&lpData->Comment()->m_lpNewProp));

			// Since lpData->data.Comment.lpNewProp was owned by an m_lpAllocParent, we don't free it directly
			if (hRes == S_OK && lpData->Comment()->m_lpNewProp)
			{
				std::wstring szTmp;
				std::wstring szAltTmp;
				SetListString(
					ulListNum,
					iItem,
					1,
					proptags::TagToString(lpData->Comment()->m_lpNewProp->ulPropTag, nullptr, false, true));
				property::parseProperty(lpData->Comment()->m_lpNewProp, &szTmp, &szAltTmp);
				SetListString(ulListNum, iItem, 2, szTmp);
				SetListString(ulListNum, iItem, 3, szAltTmp);
				return true;
			}

			return false;
		}

		void CResCommentEditor::OnOK()
		{
			CMyDialog::OnOK(); // don't need to call CEditor::OnOK

			const auto ulNewCommentProp = GetListCount(0);

			if (ulNewCommentProp && ulNewCommentProp < ULONG_MAX / sizeof(SPropValue))
			{
				auto lpNewCommentProp =
					mapi::allocate<LPSPropValue>(sizeof(SPropValue) * ulNewCommentProp, m_lpAllocParent);
				if (lpNewCommentProp)
				{
					for (ULONG paneID = 0; paneID < ulNewCommentProp; paneID++)
					{
						const auto lpData = GetListRowData(0, paneID);
						if (lpData && lpData->Comment())
						{
							if (lpData->Comment()->m_lpNewProp)
							{
								EC_H_S(mapi::MyPropCopyMore(
									&lpNewCommentProp[paneID],
									lpData->Comment()->m_lpNewProp,
									MAPIAllocateMore,
									m_lpAllocParent));
							}
							else
							{
								EC_H_S(mapi::MyPropCopyMore(
									&lpNewCommentProp[paneID],
									lpData->Comment()->m_lpOldProp,
									MAPIAllocateMore,
									m_lpAllocParent));
							}
						}
					}
					m_ulNewCommentProp = ulNewCommentProp;
					m_lpNewCommentProp = lpNewCommentProp;
				}
			}
			if (!m_lpNewCommentRes && m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes)
			{
				EC_H_S(
					mapi::HrCopyRestriction(m_lpSourceRes->res.resComment.lpRes, m_lpAllocParent, &m_lpNewCommentRes));
			}
		}

		// Note that an alloc parent is passed in to CRestrictEditor. If a parent isn't passed, we allocate one ourselves.
		// All other memory allocated in CRestrictEditor is owned by the parent and must not be freed manually
		// If we return (detach) memory to a caller, they must MAPIFreeBuffer only if they did not pass in a parent
		static std::wstring CLASS = L"CRestrictEditor"; // STRING_OK
		// Create an editor for a restriction
		// Takes LPSRestriction lpRes as input
		CRestrictEditor::CRestrictEditor(
			_In_ CWnd* pParentWnd,
			_In_opt_ LPVOID lpAllocParent,
			_In_opt_ const _SRestriction* lpRes)
			: CEditor(
				  pParentWnd,
				  IDS_RESED,
				  IDS_RESEDPROMPT,
				  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
				  IDS_ACTIONEDITRES,
				  NULL,
				  NULL)
		{
			TRACE_CONSTRUCTOR(CLASS);

			// Not copying the source restriction, but since we're modal it won't matter
			m_lpRes = lpRes;
			m_lpOutputRes = nullptr;
			m_bModified = false;
			if (!lpRes) m_bModified = true; // if we don't have a source, make sure we never look at it

			m_lpAllocParent = lpAllocParent;

			// Set up our output restriction and determine m_lpAllocParent- this will be the basis for all other allocations
			// Allocate base memory:
			m_lpOutputRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), m_lpAllocParent);
			if (!m_lpAllocParent)
			{
				m_lpAllocParent = m_lpOutputRes;
			}

			if (m_lpOutputRes)
			{
				memset(m_lpOutputRes, 0, sizeof(SRestriction));
				if (m_lpRes) m_lpOutputRes->rt = m_lpRes->rt;
			}

			SetPromptPostFix(flags::AllFlagsToString(flagRestrictionType, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RESTRICTIONTYPE, false)); // type as a number
			AddPane(viewpane::TextPane::CreateSingleLinePane(
				1, IDS_RESTRICTIONTYPE, true)); // type as a string (flagRestrictionType)
			AddPane(viewpane::TextPane::CreateMultiLinePane(
				2, IDS_RESTRICTIONTEXT, property::RestrictionToString(GetSourceRes(), nullptr), true));
		}

		CRestrictEditor::~CRestrictEditor()
		{
			TRACE_DESTRUCTOR(CLASS);
			// Only if we self allocated m_lpOutputRes and did not detach it can we free it
			if (m_lpAllocParent == m_lpOutputRes)
			{
				MAPIFreeBuffer(m_lpOutputRes);
			}
		}

		// Used to call functions which need to be called AFTER controls are created
		BOOL CRestrictEditor::OnInitDialog()
		{
			const auto bRet = CEditor::OnInitDialog();

			const auto lpSourceRes = GetSourceRes();

			if (lpSourceRes)
			{
				SetHex(0, lpSourceRes->rt);
				SetStringW(1, flags::InterpretFlags(flagRestrictionType, lpSourceRes->rt));
			}

			return bRet;
		}

		_Check_return_ const _SRestriction* CRestrictEditor::GetSourceRes() const
		{
			if (m_lpRes && !m_bModified) return m_lpRes;
			return m_lpOutputRes;
		}

		// Create our LPSRestriction from the dialog here
		void CRestrictEditor::OnOK()
		{
			CMyDialog::OnOK(); // don't need to call CEditor::OnOK
		}

		_Check_return_ LPSRestriction CRestrictEditor::DetachModifiedSRestriction()
		{
			if (!m_bModified) return nullptr;
			const auto lpRet = m_lpOutputRes;
			m_lpOutputRes = nullptr;
			return lpRet;
		}

		// CEditor::HandleChange will return the control number of what changed
		// so I can react to it if I need to
		_Check_return_ ULONG CRestrictEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			// If the restriction type changed
			if (paneID == 0)
			{
				if (!m_lpOutputRes) return paneID;
				const auto ulOldResType = m_lpOutputRes->rt;
				const auto ulNewResType = GetHex(paneID);

				if (ulOldResType == ulNewResType) return paneID;

				SetStringW(1, flags::InterpretFlags(flagRestrictionType, ulNewResType));

				m_bModified = true;
				if ((ulOldResType == RES_AND || ulOldResType == RES_OR) &&
					(ulNewResType == RES_AND || ulNewResType == RES_OR))
				{
					m_lpOutputRes->rt = ulNewResType; // just need to adjust our type and refresh
				}
				else
				{
					// Need to wipe out our current output restriction and rebuild it
					if (m_lpAllocParent == m_lpOutputRes)
					{
						// We allocated m_lpOutputRes directly, so we can and should free it before replacing the pointer
						MAPIFreeBuffer(m_lpOutputRes);
						m_lpOutputRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction));
						m_lpAllocParent = m_lpOutputRes;
					}
					else
					{
						// If the pointers are different, m_lpOutputRes was allocated with MAPIAllocateMore
						// Since m_lpOutputRes is owned by m_lpAllocParent, we don't free it directly
						m_lpOutputRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), m_lpAllocParent);
					}

					m_lpOutputRes->rt = ulNewResType;
				}

				SetStringW(2, property::RestrictionToString(m_lpOutputRes, nullptr));
			}
			return paneID;
		}

		void CRestrictEditor::OnEditAction1()
		{
			auto hRes = S_OK;
			const auto lpSourceRes = GetSourceRes();
			if (!lpSourceRes || !m_lpOutputRes) return;

			switch (lpSourceRes->rt)
			{
			case RES_COMPAREPROPS:
				hRes = WC_H(EditCompare(lpSourceRes));
				break;
				// Structures for these two types are identical
			case RES_OR:
			case RES_AND:
				hRes = WC_H(EditAndOr(lpSourceRes));
				break;
				// Structures for these two types are identical
			case RES_NOT:
			case RES_COUNT:
				hRes = WC_H(EditRestrict(lpSourceRes));
				break;
				// Structures for these two types are identical
			case RES_PROPERTY:
			case RES_CONTENT:
				hRes = WC_H(EditCombined(lpSourceRes));
				break;
			case RES_BITMASK:
				hRes = WC_H(EditBitmask(lpSourceRes));
				break;
			case RES_SIZE:
				hRes = WC_H(EditSize(lpSourceRes));
				break;
			case RES_EXIST:
				hRes = WC_H(EditExist(lpSourceRes));
				break;
			case RES_SUBRESTRICTION:
				hRes = WC_H(EditSubrestriction(lpSourceRes));
				break;
				// Structures for these two types are identical
			case RES_COMMENT:
			case RES_ANNOTATION:
				hRes = WC_H(EditComment(lpSourceRes));
				break;
			}

			if (hRes == S_OK)
			{
				m_bModified = true;
				SetStringW(2, property::RestrictionToString(m_lpOutputRes, nullptr));
			}
		}

		HRESULT CRestrictEditor::EditCompare(const _SRestriction* lpSourceRes)
		{
			CResCompareEditor MyEditor(
				this,
				lpSourceRes->res.resCompareProps.relop,
				lpSourceRes->res.resCompareProps.ulPropTag1,
				lpSourceRes->res.resCompareProps.ulPropTag2);
			if (MyEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resCompareProps.relop = MyEditor.GetHex(0);
				m_lpOutputRes->res.resCompareProps.ulPropTag1 = MyEditor.GetPropTag(2);
				m_lpOutputRes->res.resCompareProps.ulPropTag2 = MyEditor.GetPropTag(4);
				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditAndOr(const _SRestriction* lpSourceRes)
		{
			CResAndOrEditor MyResEditor(this, lpSourceRes,
										m_lpAllocParent); // pass source res into editor
			if (MyResEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resAnd.cRes = MyResEditor.GetResCount();

				const auto lpNewResArray = MyResEditor.DetachModifiedSRestrictionArray();
				if (lpNewResArray)
				{
					// I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resAnd.lpRes);
					m_lpOutputRes->res.resAnd.lpRes = lpNewResArray;
				}

				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditRestrict(const _SRestriction* lpSourceRes)
		{
			CRestrictEditor MyResEditor(
				this,
				m_lpAllocParent,
				lpSourceRes->res.resNot.lpRes); // pass source res into editor
			if (MyResEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				// Since m_lpOutputRes->res.resNot.lpRes was owned by an m_lpAllocParent, we don't free it directly
				m_lpOutputRes->res.resNot.lpRes = MyResEditor.DetachModifiedSRestriction();
				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditCombined(const _SRestriction* lpSourceRes)
		{
			auto hRes = S_OK;
			CResCombinedEditor MyEditor(
				this,
				lpSourceRes->rt,
				lpSourceRes->res.resContent.ulFuzzyLevel,
				lpSourceRes->res.resContent.ulPropTag,
				lpSourceRes->res.resContent.lpProp,
				m_lpAllocParent);
			if (!MyEditor.DisplayDialog()) return S_FALSE;

			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resContent.ulFuzzyLevel = MyEditor.GetHex(0);
			m_lpOutputRes->res.resContent.ulPropTag = MyEditor.GetPropTag(2);

			// Since m_lpOutputRes->res.resContent.lpProp was owned by an m_lpAllocParent, we don't free it directly
			m_lpOutputRes->res.resContent.lpProp = MyEditor.DetachModifiedSPropValue();

			if (!m_lpOutputRes->res.resContent.lpProp)
			{
				// Got a problem here - the relop or fuzzy level was changed, but not the property
				// Need to copy the property from the source Res to the output Res
				m_lpOutputRes->res.resContent.lpProp =
					mapi::allocate<LPSPropValue>(sizeof(SPropValue), m_lpAllocParent);

				if (m_lpOutputRes->res.resContent.lpProp)
				{
					hRes = EC_H(mapi::MyPropCopyMore(
						m_lpOutputRes->res.resContent.lpProp,
						lpSourceRes->res.resContent.lpProp,
						MAPIAllocateMore,
						m_lpAllocParent));
				}
			}

			return hRes;
		}

		HRESULT CRestrictEditor::EditBitmask(const _SRestriction* lpSourceRes)
		{
			CResBitmaskEditor MyEditor(
				this,
				lpSourceRes->res.resBitMask.relBMR,
				lpSourceRes->res.resBitMask.ulPropTag,
				lpSourceRes->res.resBitMask.ulMask);
			if (MyEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resBitMask.relBMR = MyEditor.GetHex(0);
				m_lpOutputRes->res.resBitMask.ulPropTag = MyEditor.GetPropTag(2);
				m_lpOutputRes->res.resBitMask.ulMask = MyEditor.GetHex(4);
				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditSize(const _SRestriction* lpSourceRes)
		{
			CResSizeEditor MyEditor(
				this, lpSourceRes->res.resSize.relop, lpSourceRes->res.resSize.ulPropTag, lpSourceRes->res.resSize.cb);
			if (MyEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resSize.relop = MyEditor.GetHex(0);
				m_lpOutputRes->res.resSize.ulPropTag = MyEditor.GetPropTag(2);
				m_lpOutputRes->res.resSize.cb = MyEditor.GetHex(4);
				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditExist(const _SRestriction* lpSourceRes)
		{
			CResExistEditor MyEditor(this, lpSourceRes->res.resExist.ulPropTag);
			if (MyEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resExist.ulPropTag = MyEditor.GetPropTag(0);
				m_lpOutputRes->res.resExist.ulReserved1 = 0;
				m_lpOutputRes->res.resExist.ulReserved2 = 0;
				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditSubrestriction(const _SRestriction* lpSourceRes)
		{
			CResSubResEditor MyEditor(
				this, lpSourceRes->res.resSub.ulSubObject, lpSourceRes->res.resSub.lpRes, m_lpAllocParent);
			if (MyEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resSub.ulSubObject = MyEditor.GetHex(1);

				// Since m_lpOutputRes->res.resSub.lpRes was owned by an m_lpAllocParent, we don't free it directly
				m_lpOutputRes->res.resSub.lpRes = MyEditor.DetachModifiedSRestriction();
				return S_OK;
			}

			return S_FALSE;
		}

		HRESULT CRestrictEditor::EditComment(const _SRestriction* lpSourceRes)
		{
			CResCommentEditor MyResEditor(this, lpSourceRes,
										  m_lpAllocParent); // pass source res into editor
			if (MyResEditor.DisplayDialog())
			{
				m_lpOutputRes->rt = lpSourceRes->rt;

				// Since m_lpOutputRes->res.resComment.lpRes was owned by an m_lpAllocParent, we don't free it directly
				m_lpOutputRes->res.resComment.lpRes = MyResEditor.DetachModifiedSRestriction();

				// Since m_lpOutputRes->res.resComment.lpProp was owned by an m_lpAllocParent, we don't free it directly
				m_lpOutputRes->res.resComment.lpProp = MyResEditor.DetachModifiedSPropValue();
				if (m_lpOutputRes->res.resComment.lpProp)
				{
					m_lpOutputRes->res.resComment.cValues = MyResEditor.GetSPropValueCount();
				}

				return S_OK;
			}

			return S_FALSE;
		}

		// Note that no alloc parent is passed in to CCriteriaEditor. So we're completely responsible for freeing any memory we allocate.
		// If we return (detach) memory to a caller, they must MAPIFreeBuffer
		static std::wstring CRITERIACLASS = L"CCriteriaEditor"; // STRING_OK
#define LISTNUM 4
		CCriteriaEditor::CCriteriaEditor(
			_In_ CWnd* pParentWnd,
			_In_ const _SRestriction* lpRes,
			_In_ LPENTRYLIST lpEntryList,
			ULONG ulSearchState)
			: CEditor(
				  pParentWnd,
				  IDS_CRITERIAEDITOR,
				  IDS_CRITERIAEDITORPROMPT,
				  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
				  IDS_ACTIONEDITRES,
				  NULL,
				  NULL)
		{
			TRACE_CONSTRUCTOR(CRITERIACLASS);

			m_lpSourceRes = lpRes;
			m_lpNewRes = nullptr;
			m_lpSourceEntryList = lpEntryList;

			m_lpNewEntryList = mapi::allocate<SBinaryArray*>(sizeof(SBinaryArray));

			m_ulNewSearchFlags = NULL;

			SetPromptPostFix(flags::AllFlagsToString(flagSearchFlag, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SEARCHSTATE, true));
			SetHex(0, ulSearchState);
			const auto szFlags = flags::InterpretFlags(flagSearchState, ulSearchState);
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_SEARCHSTATE, szFlags, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_SEARCHFLAGS, false));
			SetHex(2, 0);
			AddPane(viewpane::TextPane::CreateSingleLinePane(3, IDS_SEARCHFLAGS, true));
			AddPane(
				viewpane::ListPane::CreateCollapsibleListPane(4, IDS_EIDLIST, false, false, ListEditCallBack(this)));
			SetListID(4);
			AddPane(viewpane::TextPane::CreateMultiLinePane(
				5, IDS_RESTRICTIONTEXT, property::RestrictionToString(m_lpSourceRes, nullptr), true));
		}

		CCriteriaEditor::~CCriteriaEditor()
		{
			// If these structures weren't detached, we need to free them
			MAPIFreeBuffer(m_lpNewEntryList);
			MAPIFreeBuffer(m_lpNewRes);
		}

		// Used to call functions which need to be called AFTER controls are created
		BOOL CCriteriaEditor::OnInitDialog()
		{
			const auto bRet = CEditor::OnInitDialog();

			InitListFromEntryList(LISTNUM, m_lpSourceEntryList);

			UpdateButtons();

			return bRet;
		}

		_Check_return_ const _SRestriction* CCriteriaEditor::GetSourceRes() const
		{
			if (m_lpNewRes) return m_lpNewRes;
			return m_lpSourceRes;
		}

		_Check_return_ ULONG CCriteriaEditor::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == 2)
			{
				SetStringW(3, flags::InterpretFlags(flagSearchFlag, GetHex(paneID)));
			}

			return paneID;
		}

		// Whoever gets this MUST MAPIFreeBuffer
		_Check_return_ LPSRestriction CCriteriaEditor::DetachModifiedSRestriction()
		{
			const auto lpRet = m_lpNewRes;
			m_lpNewRes = nullptr;
			return lpRet;
		}

		// Whoever gets this MUST MAPIFreeBuffer
		_Check_return_ LPENTRYLIST CCriteriaEditor::DetachModifiedEntryList()
		{
			const auto lpRet = m_lpNewEntryList;
			m_lpNewEntryList = nullptr;
			return lpRet;
		}

		_Check_return_ ULONG CCriteriaEditor::GetSearchFlags() const { return m_ulNewSearchFlags; }

		void CCriteriaEditor::InitListFromEntryList(ULONG ulListNum, _In_ const SBinaryArray* lpEntryList) const
		{
			ClearList(ulListNum);

			InsertColumn(ulListNum, 0, IDS_SHARP);
			InsertColumn(ulListNum, 1, IDS_CB);
			InsertColumn(ulListNum, 2, IDS_BINARY);
			InsertColumn(ulListNum, 3, IDS_TEXTVIEW);

			if (lpEntryList)
			{
				for (ULONG iRow = 0; iRow < lpEntryList->cValues; iRow++)
				{
					auto lpData = InsertListRow(ulListNum, iRow, std::to_wstring(iRow));
					if (lpData)
					{
						lpData->InitializeBinary(&lpEntryList->lpbin[iRow]);
					}

					SetListString(ulListNum, iRow, 1, std::to_wstring(lpEntryList->lpbin[iRow].cb));
					SetListString(ulListNum, iRow, 2, strings::BinToHexString(&lpEntryList->lpbin[iRow], false));
					SetListString(ulListNum, iRow, 3, strings::BinToTextString(&lpEntryList->lpbin[iRow], true));
					if (lpData) lpData->bItemFullyLoaded = true;
				}
			}

			ResizeList(ulListNum, false);
		}

		void CCriteriaEditor::OnEditAction1()
		{
			const auto lpSourceRes = GetSourceRes();

			CRestrictEditor MyResEditor(this, nullptr,
										lpSourceRes); // pass source res into editor
			if (MyResEditor.DisplayDialog())
			{
				const auto lpModRes = MyResEditor.DetachModifiedSRestriction();
				if (lpModRes)
				{
					// We didn't pass an alloc parent to CRestrictEditor, so we must free what came back
					MAPIFreeBuffer(m_lpNewRes);
					m_lpNewRes = lpModRes;
					SetStringW(5, property::RestrictionToString(m_lpNewRes, nullptr));
				}
			}
		}

		_Check_return_ bool
		CCriteriaEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData)
		{
			if (!lpData) return false;

			if (!lpData->Binary())
			{
				lpData->InitializeBinary(nullptr);
			}

			CEditor BinEdit(this, IDS_EIDEDITOR, IDS_EIDEDITORPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			auto lpSourcebin = LPSBinary{};
			if (lpData->Binary()->m_OldBin.lpb)
			{
				lpSourcebin = &lpData->Binary()->m_OldBin;
			}
			else
			{
				lpSourcebin = &lpData->Binary()->m_NewBin;
			}

			BinEdit.AddPane(viewpane::TextPane::CreateSingleLinePane(
				0, IDS_EID, strings::BinToHexString(lpSourcebin, false), false));

			if (BinEdit.DisplayDialog())
			{
				auto bin = strings::HexStringToBin(BinEdit.GetStringW(0));
				lpData->Binary()->m_NewBin.lpb = mapi::ByteVectorToMAPI(bin, m_lpNewEntryList);
				if (lpData->Binary()->m_NewBin.lpb)
				{
					lpData->Binary()->m_NewBin.cb = static_cast<ULONG>(bin.size());
					const auto szTmp = std::to_wstring(lpData->Binary()->m_NewBin.cb);
					SetListString(ulListNum, iItem, 1, szTmp);
					SetListString(ulListNum, iItem, 2, strings::BinToHexString(&lpData->Binary()->m_NewBin, false));
					SetListString(ulListNum, iItem, 3, strings::BinToTextString(&lpData->Binary()->m_NewBin, true));
					return true;
				}
			}

			return false;
		}

		void CCriteriaEditor::OnOK()
		{
			CMyDialog::OnOK(); // don't need to call CEditor::OnOK
			const auto ulValues = GetListCount(LISTNUM);

			if (m_lpNewEntryList && ulValues < ULONG_MAX / sizeof(SBinary))
			{
				m_lpNewEntryList->cValues = ulValues;
				m_lpNewEntryList->lpbin =
					mapi::allocate<LPSBinary>(m_lpNewEntryList->cValues * sizeof(SBinary), m_lpNewEntryList);

				for (ULONG paneID = 0; paneID < m_lpNewEntryList->cValues; paneID++)
				{
					const auto lpData = GetListRowData(LISTNUM, paneID);
					if (lpData && lpData->Binary())
					{
						if (lpData->Binary()->m_NewBin.lpb)
						{
							m_lpNewEntryList->lpbin[paneID].cb = lpData->Binary()->m_NewBin.cb;
							m_lpNewEntryList->lpbin[paneID].lpb = lpData->Binary()->m_NewBin.lpb;
							// clean out the source
							lpData->Binary()->m_OldBin.lpb = nullptr;
						}
						else
						{
							m_lpNewEntryList->lpbin[paneID].cb = lpData->Binary()->m_OldBin.cb;
							m_lpNewEntryList->lpbin[paneID].lpb =
								mapi::allocate<LPBYTE>(m_lpNewEntryList->lpbin[paneID].cb, m_lpNewEntryList);

							memcpy(
								m_lpNewEntryList->lpbin[paneID].lpb,
								lpData->Binary()->m_OldBin.lpb,
								m_lpNewEntryList->lpbin[paneID].cb);
						}
					}
				}
			}

			if (!m_lpNewRes && m_lpSourceRes)
			{
				EC_H_S(mapi::HrCopyRestriction(m_lpSourceRes, NULL, &m_lpNewRes));
			}

			m_ulNewSearchFlags = GetHex(2);
		}
	} // namespace editor
} // namespace dialog