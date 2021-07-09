#pragma once

typedef struct TypeInfoNode
{
    TypeInfoNode();
    TypeInfoNode(LPTYPEINFO pTypeInfo, MEMBERID memberID = MEMBERID_NIL);
    ~TypeInfoNode();

    CComPtr<ITypeInfo> pTypeInfo;
    MEMBERID memberID;
}* LPTYPEINFONODE;
