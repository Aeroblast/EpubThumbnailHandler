/********************************  ********************************\
 修改自RecipeThumbnailProvider.cpp

\*******************************************************************************/

#include "EpubThumbnailProvider.h"
#include <Shlwapi.h>
#include <Wincrypt.h>   // For CryptStringToBinary.
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "msxml6.lib")
#include<regex>
//debug
#include<string>
#include<fstream>
#include <iostream>
#ifdef _DEBUG
#define DEBUGLOG(str) {std::wofstream file;file.open("E:\\log.txt", std::ios::app);file << str<< "\r\n";}
#else
#define DEBUGLOG(str) ;
#endif // DEBUG


//helpers
HRESULT OPFHandle(char* text, BSTR* coverpath);
HRESULT MetaCheck(std::string t, std::string *id);
HRESULT  GetOPFPath(char*text, BSTR* opf_path);
HRESULT GetCoverPath(HZIP zip, BSTR*path);
HRESULT ItemCheck(std::string t, std::string id, std::string*href);
BSTR CombineHref(BSTR referer, BSTR href);



extern HINSTANCE g_hInst;
extern long g_cDllRef;


EpubThumbnailProvider::EpubThumbnailProvider() : m_cRef(1), m_pStream(NULL)
{
    InterlockedIncrement(&g_cDllRef);
}


EpubThumbnailProvider::~EpubThumbnailProvider()
{
    InterlockedDecrement(&g_cDllRef);
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP EpubThumbnailProvider::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(EpubThumbnailProvider, IThumbnailProvider),
        QITABENT(EpubThumbnailProvider, IInitializeWithStream), 
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) EpubThumbnailProvider::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) EpubThumbnailProvider::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IInitializeWithStream

// Initializes the thumbnail handler with a stream.
IFACEMETHODIMP EpubThumbnailProvider::Initialize(IStream *pStream, DWORD grfMode)
{
    // A handler instance should be initialized only once in its lifetime. 
    HRESULT hr = HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
	
    if (m_pStream == NULL)
    {
        // Take a reference to the stream if it has not been initialized yet.
        hr = pStream->QueryInterface(&m_pStream);
    }
    return hr;
}

#pragma endregion


#pragma region IThumbnailProvider

// Gets a thumbnail image and alpha type. The GetThumbnail is called with the 
// largest desired size of the image, in pixels. Although the parameter is 
// called cx, this is used as the maximum size of both the x and y dimensions. 
// If the retrieved thumbnail is not square, then the longer axis is limited 
// by cx and the aspect ratio of the original image respected. On exit, 
// GetThumbnail provides a handle to the retrieved image. It also provides a 
// value that indicates the color format of the image and whether it has 
// valid alpha information.
IFACEMETHODIMP EpubThumbnailProvider::GetThumbnail(UINT cx, HBITMAP *phbmp, 
    WTS_ALPHATYPE *pdwAlpha)
{
	HRESULT hr;
	ULONG n = 1;
	ULONG actRead;
	ULARGE_INTEGER fileSize;
	DEBUGLOG("run");
	hr=IStream_Size(m_pStream,&fileSize); if (hr != S_OK) { DEBUGLOG("Error at IStream_Size"); return E_ABORT; }
	
	byte* buffer = (byte*)CoTaskMemAlloc(fileSize.QuadPart); if (!buffer) { DEBUGLOG("Error at malloc buffer");  return E_ABORT; }
	hr = m_pStream->Read(buffer, (LONG)fileSize.QuadPart+1, &actRead);if (hr != S_FALSE) {CoTaskMemFree(buffer); DEBUGLOG("Error at Read Stream"); return E_ABORT;}

	//得到了全部数据
	
	//加载zip
	HZIP zip= OpenZip(buffer,actRead,0);if(!zip) { DEBUGLOG("Error at Load zip"); CoTaskMemFree(buffer); return E_ABORT; }

	//查找用
	int index = -1; ZIPENTRY entry;

	//假定路径，如果中了就省了不少事
	hr=FindZipItem(zip,L"OEBPS/Images/cover.jpg",true,&index,&entry);
	if(index!=-1)
	{
		byte*image = (byte*)CoTaskMemAlloc(entry.unc_size); if (!image) { CoTaskMemFree(buffer); return E_ABORT; }
		hr=UnzipItem(zip, index, image, entry.unc_size);
		if (hr == ZR_OK) 
		{
			IStream *pImageStream = SHCreateMemStream(image, entry.unc_size);
			hr=WICCreate32bppHBITMAP(pImageStream, phbmp, pdwAlpha);
			pImageStream->Release();
			CoTaskMemFree(image);
			CloseZip(zip);CoTaskMemFree(buffer);
			return hr;
		}
	}
	
	//get root-file
	{
		BSTR coverpath;
		hr=GetCoverPath(zip,&coverpath);
		if (SUCCEEDED(hr)) 
		{
			hr = FindZipItem(zip,coverpath, false, &index, &entry);
			if (index != -1)
			{
				byte*image = (byte*)CoTaskMemAlloc(entry.unc_size); if (!image) { CoTaskMemFree(buffer); return E_ABORT; }
				hr = UnzipItem(zip, index, image, entry.unc_size);
				if (hr == ZR_OK)
				{
					IStream *pImageStream = SHCreateMemStream(image, entry.unc_size);
					hr = WICCreate32bppHBITMAP(pImageStream, phbmp, pdwAlpha);
					pImageStream->Release();
					CoTaskMemFree(image);
					CloseZip(zip);
				}
			}
			SysFreeString(coverpath);
		}

	}
	

	CoTaskMemFree(buffer);
    return hr;
}

#pragma endregion


#pragma region Helper Functions



HRESULT GetCoverPath(HZIP zip, BSTR*path)
{

	int index = -1; ZIPENTRY entry;
	HRESULT hr = FindZipItem(zip, L"META-INF/container.xml", false, &index, &entry);
	if (index == -1) { return E_FAIL; }
	char* xml = (char*)CoTaskMemAlloc(entry.unc_size);
	if (xml)
	{
		hr = UnzipItem(zip, index, xml, entry.unc_size);
		if (SUCCEEDED(hr)) //获得container.xml
		{
			DEBUGLOG("Got container.xml");

			BSTR opf_path = NULL;
			hr = GetOPFPath(xml, &opf_path);//解析container.xml
			if (SUCCEEDED(hr))//得到opf路径
			{
				DEBUGLOG(opf_path);
				hr = FindZipItem(zip, opf_path, false, &index, &entry);
				if (index != -1)
				{
					char* opf_data = (char*)CoTaskMemAlloc(entry.unc_size);
					if (opf_data)
					{
						hr = UnzipItem(zip, index, opf_data, entry.unc_size);
						if (SUCCEEDED(hr))
						{
							DEBUGLOG("Got opf data");
							BSTR cover_path = NULL;
							hr = OPFHandle(opf_data, &cover_path);
							*path = CombineHref(opf_path, cover_path);
							SysFreeString(cover_path);
						}

					}
					else { hr = E_FAIL; }
					CoTaskMemFree(opf_data);
				}
				else
				{
					hr = E_FAIL;
				}

			}
			SysFreeString(opf_path);

		}
		CoTaskMemFree(xml);

	}

	return hr;
}

HRESULT  GetOPFPath(char*text, BSTR* opf_path)
{
	std::regex rx("full-path=\"(.*?)\"");
	std::string t = text;
	std::smatch sm;
	std::regex_search(t, sm, rx);
	if (sm.length() > 1)
	{
		std::string r = sm[1].str();
		const size_t strsize = r.length() + 1;
		*opf_path = SysAllocStringLen(NULL, (UINT)strsize);
		mbstowcs(*opf_path, r.c_str(), strsize);
		return S_OK;
	}
	else
	{
		return E_FAIL;
	}
}

HRESULT OPFHandle(char* text, BSTR* cover_path)
{
	std::regex reg_meta("<meta .*>");
	std::regex reg_item("<item .*?>");
	std::string t = text;
	HRESULT hr;
	std::string id;

	std::string temp = t;
	std::smatch m;
	while (std::regex_search(temp, m, reg_meta)) {
		hr = MetaCheck(m[0].str(), &id);
		if (SUCCEEDED(hr))
		{
			temp = t;
			std::smatch m;
			while (std::regex_search(temp, m, reg_item))
			{
				std::string href;
				hr = ItemCheck(m[0].str(), id, &href);
				if (SUCCEEDED(hr))
				{
					*cover_path = SysAllocStringLen(NULL, (UINT)href.length());
					mbstowcs(*cover_path, href.c_str(), href.length());
					DEBUGLOG(cover_path);
					return S_OK;
					//break;
				}
				temp = m.suffix().str();
			}
			//break;
		}
		temp = m.suffix().str();
	}
	return E_FAIL;
}

HRESULT MetaCheck(std::string t, std::string *id)
{
	char name[] = "name=\"cover\"";
	char content[] = "content=\"";
	if (t.find(name) != t.npos)
	{
		size_t pos = t.find(content);
		if (pos != t.npos)
		{
			pos += strlen(content);
			size_t end = t.find('"', pos);
			if (end != t.npos)
			{
				*id = t.substr(pos, end - pos);
				return S_OK;
			}
		}
	}
	return E_FAIL;
}
HRESULT ItemCheck(std::string t, std::string id, std::string*href)
{
	std::string idvalue = "id=\"" + id + "\"";
	char content[] = "href=\"";
	if (t.find(idvalue) != t.npos)
	{
		size_t pos = t.find(content);
		if (pos != t.npos)
		{
			pos += strlen(content);
			size_t end = t.find('"', pos);
			if (end != t.npos)
			{
				*href = t.substr(pos, end - pos);
				return S_OK;
			}
		}
	}
	return E_FAIL;
}
BSTR CombineHref(BSTR referer, BSTR href)
{

	UINT len = SysStringLen(referer);
	UINT len2 = SysStringLen(href);
	int i = len - 1;
	for (; i >= 0; i--)
	{
		if (referer[i] == L'/')break;
	}

	if (i == 0)
	{
		BSTR result;
		SHStrDup(href, &result);
		return result;
	}
	else
	{
		BSTR result = SysAllocStringLen(NULL, i + 1 + len2 + 1);
		memcpy(result, referer, (i + 1) * sizeof(OLECHAR));
		memcpy(result + (i + 1), href, len2 * sizeof(OLECHAR));
		result[i + 1 + len2] = 0;
		return result;
	}

}


// The function reads the Source attribute of the Recipe/Attachments/Picture 
// node in the specified XML document, and gets the base64-encoded string 
// value which represents the thumbnail image. The cx paramter specifies the 
// desired size of the image. For simplicity, the sample omits the cx 
// paramter and provides only one image for all situations.
HRESULT EpubThumbnailProvider::GetBase64EncodedImageString(
    IXMLDOMDocument *pXMLDoc, UINT /*cx*/, PWSTR *ppszResult)
{
    *ppszResult = NULL;
    HRESULT hr;

    // Get the node Recipe/Attachments/Picture in the XML document.
    BSTR bstrQuery = SysAllocString(L"Recipe/Attachments/Picture");
    hr = bstrQuery ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        IXMLDOMNode *pXMLNode;
        hr = pXMLDoc->selectSingleNode(bstrQuery, &pXMLNode);
        if (SUCCEEDED(hr))
        {
            IXMLDOMElement *pXMLElement;
            hr = pXMLNode->QueryInterface(&pXMLElement);
            if (SUCCEEDED(hr))
            {
                // Read the "Source" attribute of the XML node.
                BSTR bstrAttribute = SysAllocString(L"Source");
                hr = bstrAttribute ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    VARIANT varValue;
                    hr = pXMLElement->getAttribute(bstrAttribute, &varValue);
                    if (SUCCEEDED(hr))
                    {
                        // Output the "Source" string attribute.
                        if ((varValue.vt == VT_BSTR) && varValue.bstrVal && 
                            varValue.bstrVal[0] /* not an empty string */)
                        {
                            hr = SHStrDup(varValue.bstrVal, ppszResult);
                        }
                        else
                        {
                            hr = E_FAIL;
                        }
                        VariantClear(&varValue);
                    }
                    SysFreeString(bstrAttribute);
                }
                pXMLElement->Release();
            }
            pXMLNode->Release();
        }
        SysFreeString(bstrQuery);
    }

    return hr;
}


HRESULT EpubThumbnailProvider::ConvertBitmapSourceTo32bppHBITMAP(
    IWICBitmapSource *pBitmapSource, IWICImagingFactory *pImagingFactory, 
    HBITMAP *phbmp)
{
    *phbmp = NULL;

    IWICBitmapSource *pBitmapSourceConverted = NULL;
    WICPixelFormatGUID guidPixelFormatSource;
    HRESULT hr = pBitmapSource->GetPixelFormat(&guidPixelFormatSource);

    if (SUCCEEDED(hr) && (guidPixelFormatSource != GUID_WICPixelFormat32bppBGRA))
    {
        IWICFormatConverter *pFormatConverter;
        hr = pImagingFactory->CreateFormatConverter(&pFormatConverter);
        if (SUCCEEDED(hr))
        {
            // Create the appropriate pixel format converter.
            hr = pFormatConverter->Initialize(pBitmapSource, 
                GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, NULL, 
                0, WICBitmapPaletteTypeCustom);
            if (SUCCEEDED(hr))
            {
                hr = pFormatConverter->QueryInterface(&pBitmapSourceConverted);
            }
            pFormatConverter->Release();
        }
    }
    else
    {
        // No conversion is necessary.
        hr = pBitmapSource->QueryInterface(&pBitmapSourceConverted);
    }

    if (SUCCEEDED(hr))
    {
        UINT nWidth, nHeight;
        hr = pBitmapSourceConverted->GetSize(&nWidth, &nHeight);
        if (SUCCEEDED(hr))
        {
            BITMAPINFO bmi = { sizeof(bmi.bmiHeader) };
            bmi.bmiHeader.biWidth = nWidth;
            bmi.bmiHeader.biHeight = -static_cast<LONG>(nHeight);
            bmi.bmiHeader.biPlanes = 1;
            bmi.bmiHeader.biBitCount = 32;
            bmi.bmiHeader.biCompression = BI_RGB;

            BYTE *pBits;
            HBITMAP hbmp = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, 
                reinterpret_cast<void **>(&pBits), NULL, 0);
            hr = hbmp ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                WICRect rect = {0, 0, (INT)nWidth, (INT)nHeight};

                // Convert the pixels and store them in the HBITMAP.  
                // Note: the name of the function is a little misleading - 
                // we're not doing any extraneous copying here.  CopyPixels 
                // is actually converting the image into the given buffer.
                hr = pBitmapSourceConverted->CopyPixels(&rect, nWidth * 4, 
                    nWidth * nHeight * 4, pBits);
                if (SUCCEEDED(hr))
                {
                    *phbmp = hbmp;
                }
                else
                {
                    DeleteObject(hbmp);
                }
            }
        }
        pBitmapSourceConverted->Release();
    }
    return hr;
}


HRESULT EpubThumbnailProvider::WICCreate32bppHBITMAP(IStream *pstm, 
    HBITMAP *phbmp, WTS_ALPHATYPE *pdwAlpha)
{
    *phbmp = NULL;

    // Create the COM imaging factory.
    IWICImagingFactory *pImagingFactory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, NULL, 
        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pImagingFactory));
    if (SUCCEEDED(hr))
    {
        // Create an appropriate decoder.
        IWICBitmapDecoder *pDecoder;
        hr = pImagingFactory->CreateDecoderFromStream(pstm, 
            &GUID_VendorMicrosoft, WICDecodeMetadataCacheOnDemand, &pDecoder);
        if (SUCCEEDED(hr))
        {
            IWICBitmapFrameDecode *pBitmapFrameDecode;
            hr = pDecoder->GetFrame(0, &pBitmapFrameDecode);
            if (SUCCEEDED(hr))
            {
                hr = ConvertBitmapSourceTo32bppHBITMAP(pBitmapFrameDecode, 
                    pImagingFactory, phbmp);
                if (SUCCEEDED(hr))
                {
                    *pdwAlpha = WTSAT_ARGB;
                }
                pBitmapFrameDecode->Release();
            }
            pDecoder->Release();
        }
        pImagingFactory->Release();
    }
    return hr;
}

#pragma endregion