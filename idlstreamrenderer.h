#pragma once

#include "textstream.h"

class IDLStreamRenderer
{
public:
    IDLStreamRenderer();

    void Parse(TextStream& stream);
    void Render(CDCHandle dc);

    CSize GetCharSize() const;
    CSize GetDocSize() const;

private:
    int m_cxChar = 0;
    int m_cyChar = 0;
    CRect m_rc;
    CFont m_font;
    CDC m_memDC;
    CBitmap m_bitmap;
};
