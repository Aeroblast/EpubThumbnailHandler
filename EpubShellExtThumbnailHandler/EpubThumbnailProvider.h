
#pragma once

#include <windows.h>
#include <thumbcache.h>     // For IThumbnailProvider
#include <wincodec.h>       // Windows Imaging Codecs
#include "unzip.h"

#pragma comment(lib, "windowscodecs.lib")


class EpubThumbnailProvider : 
    public IInitializeWithStream, 
    public IThumbnailProvider
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode);

    // IThumbnailProvider
    IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pdwAlpha);

    EpubThumbnailProvider();

protected:
    ~EpubThumbnailProvider();

private:
    // Reference count of component.
    long m_cRef;

    // Provided during initialization.
    IStream *m_pStream;


    HRESULT LoadXMLDocument(IXMLDOMDocument **ppXMLDoc,IStream*pStream);

    HRESULT GetBase64EncodedImageString(
        IXMLDOMDocument *pXMLDoc, 
        UINT cx, 
        PWSTR *ppszResult);


    HRESULT ConvertBitmapSourceTo32bppHBITMAP(
        IWICBitmapSource *pBitmapSource, 
        IWICImagingFactory *pImagingFactory, 
        HBITMAP *phbmp);

    HRESULT WICCreate32bppHBITMAP(IStream *pstm, HBITMAP *phbmp, 
        WTS_ALPHATYPE *pdwAlpha);
};