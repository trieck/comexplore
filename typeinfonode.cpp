#include "stdafx.h"
#include "typeinfonode.h"

TypeInfoNode::TypeInfoNode()
    : type(TypeInfoType::T_UNKNOWN), pTypeInfo(nullptr), memberID(MEMBERID_NIL)
{
}

TypeInfoNode::TypeInfoNode(TypeInfoType type, LPTYPEINFO pTypeInfo, MEMBERID memberID)
    : type(type), pTypeInfo(pTypeInfo), memberID(memberID)
{
}

TypeInfoNode::~TypeInfoNode()
{
    if (pTypeInfo != nullptr) {
        auto* p = pTypeInfo.Detach();
        auto result = p->Release();
        ATLTRACE("(0x%p) ITypeInfo::Release() == %d.\n", p, result);
    }
}
