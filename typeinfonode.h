#pragma once

/////////////////////////////////////////////////////////////////////////////
enum class TypeInfoType
{
    T_UNKNOWN = 0,
    T_ENUM,
    T_RECORD,
    T_MODULE,
    T_INTERFACE,
    T_DISPATCH,
    T_FUNCTION,
    T_COCLASS,
    T_ALIAS,
    T_UNION,
    T_VAR,
    T_CONST
};

/////////////////////////////////////////////////////////////////////////////
typedef struct TypeInfoNode
{
    TypeInfoNode();
    TypeInfoNode(TypeInfoType type, LPTYPEINFO pTypeInfo, MEMBERID memberID = MEMBERID_NIL);
    ~TypeInfoNode();

    TypeInfoType type;
    CComPtr<ITypeInfo> pTypeInfo;
    MEMBERID memberID;
}* LPTYPEINFONODE;
