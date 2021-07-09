#include "stdafx.h"
#include "typeinfonode.h"

TypeInfoNode::TypeInfoNode()
    : pTypeInfo(nullptr), memberID(MEMBERID_NIL)
{
}

TypeInfoNode::TypeInfoNode(LPTYPEINFO pTypeInfo, MEMBERID memberID)
    : pTypeInfo(pTypeInfo), memberID(memberID)
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
