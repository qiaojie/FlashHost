#ifndef GRAPHICSDC_H
#define GRAPHICSDC_H
#ifndef __ATLGDI_H__
    #error graphicsdc.h requires atlgdi.h to be included first
#endif

class GraphicsDC : public CDC
{
public:
    // Data members
    SIZE m_szImage;
    CBitmap m_bmp;
    HDC     m_hOrigDC;
    HBITMAP m_hBmpOld;

    // Constructor/destructor
    GraphicsDC(HDC hDC, int cx, int cy) : m_hOrigDC(NULL), m_hBmpOld(NULL)
    {
        InitGraphic(hDC, cx, cy);
    }

    ~GraphicsDC()
    {
        DeleteGraphic();
    }

    bool InitGraphic(HDC hDC, int cx, int cy)
    {
        DeleteGraphic();
        if ( cx > 0 && cy > 0 )
        {
            m_szImage.cx = cx;
            m_szImage.cy = cy;
            CreateCompatibleDC(hDC);
            ATLASSERT(m_hDC != NULL);
            m_bmp.CreateCompatibleBitmap(hDC, m_szImage.cx, m_szImage.cy);
            ATLASSERT(m_bmp.m_hBitmap != NULL);
            m_hBmpOld = SelectBitmap(m_bmp);
            m_hOrigDC = hDC;
        }
        return ( m_hDC != NULL );
    }

    void DeleteGraphic()
    {
        if (m_hDC != NULL)
        {
            SelectBitmap(m_hBmpOld);
            ::DeleteDC(Detach());
            m_bmp.DeleteObject();
        }
        m_szImage.cx = m_szImage.cy = 0;
        m_hOrigDC = NULL;
    }
};

#endif // GRAPHICSDC_H