#pragma once

template <int count = 1>
class AutoBstrArray
{
public:
    AutoBstrArray();
    ~AutoBstrArray();

    int size() const;
    explicit operator LPBSTR();

    BSTR& operator[](int index);

private:
    void Free();
    BSTR m_str[count];
};

template <int count>
AutoBstrArray<count>::AutoBstrArray(): m_str{}
{
}

template <int count>
AutoBstrArray<count>::~AutoBstrArray()
{
    Free();
}

template <int count>
int AutoBstrArray<count>::size() const
{
    return count;
}

template <int count>
AutoBstrArray<count>::operator LPBSTR()
{
    return &m_str[0];
}

template <int count>
BSTR& AutoBstrArray<count>::operator[](int index)
{
    ATLASSERT(index >= 0 && index < count);
    return m_str[index];
}

template <int count>
void AutoBstrArray<count>::Free()
{
    for (auto i = 0; i < count; ++i) {
        if (m_str[i] != nullptr) {
            SysFreeString(m_str[i]);
            m_str[i] = nullptr;
        }
    }
}
