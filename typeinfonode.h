#pragma once

typedef struct TypeInfoNode
{
    TypeInfoNode();
    TypeInfoNode(LPTYPEINFO pTypeInfo, MEMBERID memberID = MEMBERID_NIL);
    ~TypeInfoNode();

    LPTYPEINFO pTypeInfo;
    MEMBERID memberID;
}* LPTYPEINFONODE;
