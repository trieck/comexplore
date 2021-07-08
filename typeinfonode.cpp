#include "stdafx.h"
#include "typeinfonode.h"

TypeInfoNode::TypeInfoNode()
    : pTypeInfo(nullptr), memberID(MEMBERID_NIL)
{
}

TypeInfoNode::TypeInfoNode(LPTYPEINFO pTypeInfo, MEMBERID memberID)
    : pTypeInfo(pTypeInfo), memberID(memberID)
{
    pTypeInfo->AddRef();
}

TypeInfoNode::~TypeInfoNode()
{
    if (pTypeInfo != nullptr) {
        auto* p = pTypeInfo;
        auto result = pTypeInfo->Release();
        ATLTRACE("(0x%p) ITypeInfo::Release() == %d.\n", p, result);
    }
}
