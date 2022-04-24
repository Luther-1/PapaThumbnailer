// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

#include <shlwapi.h>
#include <Wincrypt.h>   // For CryptStringToBinary.
#include <thumbcache.h> // For IThumbnailProvider.
#include <wincodec.h>   // Windows Imaging Codecs
#include <msxml6.h>
#include <new>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "msxml6.lib")

// this thumbnail provider implements IInitializeWithStream to enable being hosted
// in an isolated process for robustness

class CPapaThumbProvider : public IInitializeWithStream,
                             public IThumbnailProvider
{
public:
    CPapaThumbProvider() : _cRef(1), _pStream(NULL)
    {
    }

    virtual ~CPapaThumbProvider()
    {
        if (_pStream)
        {
            _pStream->Release();
        }
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CPapaThumbProvider, IInitializeWithStream),
            QITABENT(CPapaThumbProvider, IThumbnailProvider),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG cRef = InterlockedDecrement(&_cRef);
        if (!cRef)
        {
            delete this;
        }
        return cRef;
    }

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode);

    // IThumbnailProvider
    IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pdwAlpha);

private:

    long _cRef;
    IStream *_pStream;     // provided during initialization.
    VOID DxtDecodeColourMap(BYTE*, UINT, BYTE[4][3]);
    VOID DxtDecodeAlphaMap(BYTE*, UINT, BYTE[16]);
    VOID DecodeTexture(BYTE*, USHORT, USHORT, BYTE, BYTE*);

};

HRESULT CPapaThumbProvider_CreateInstance(REFIID riid, void **ppv)
{
    CPapaThumbProvider *pNew = new (std::nothrow) CPapaThumbProvider();
    HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pNew->QueryInterface(riid, ppv);
        pNew->Release();
    }
    return hr;
}

// IInitializeWithStream
IFACEMETHODIMP CPapaThumbProvider::Initialize(IStream *pStream, DWORD)
{
    HRESULT hr = E_UNEXPECTED;  // can only be inited once
    if (_pStream == NULL)
    {
        // take a reference to the stream if we have not been inited yet
        hr = pStream->QueryInterface(&_pStream);
    }
    return hr;
}

VOID CPapaThumbProvider::DxtDecodeColourMap(BYTE* data, UINT dataLoc, BYTE colours[4][3]) { // [[R,G,B] * 4]
    UINT colour0 = (data[dataLoc + 0]) | (data[dataLoc + 1] << 8);
    UINT colour1 = (data[dataLoc + 2]) | (data[dataLoc + 3] << 8);

    colours[0][0] = (colour0 >> 8) & 0b11111000;
    colours[0][1] = (colour0 >> 3) & 0b11111100;
    colours[0][2] = (colour0 << 3) & 0b11111000;

    colours[1][0] = (colour1 >> 8) & 0b11111000;
    colours[1][1] = (colour1 >> 3) & 0b11111100;
    colours[1][2] = (colour1 << 3) & 0b11111000;

    if (colour0 > colour1) {
        colours[2][0] = (BYTE)((2 * (UINT)colours[0][0] + (UINT)colours[1][0]) / 3.0);
        colours[2][1] = (BYTE)((2 * (UINT)colours[0][1] + colours[1][1]) / 3.0);
        colours[2][2] = (BYTE)((2 * (UINT)colours[0][2] + colours[1][2]) / 3.0);

        colours[3][0] = (BYTE)(((UINT)colours[0][0] + 2 * colours[1][0]) / 3.0);
        colours[3][1] = (BYTE)(((UINT)colours[0][1] + 2 * colours[1][1]) / 3.0);
        colours[3][2] = (BYTE)(((UINT)colours[0][2] + 2 * colours[1][2]) / 3.0);
    }
    else {
        colours[2][0] = (BYTE)(((UINT)colours[0][0] + colours[1][0]) / 2.0);
        colours[2][1] = (BYTE)(((UINT)colours[0][1] + colours[1][1]) / 2.0);
        colours[2][2] = (BYTE)(((UINT)colours[0][2] + colours[1][2]) / 2.0);

        colours[3][0] = 0;
        colours[3][1] = 0;
        colours[3][2] = 0;

    }
}

VOID CPapaThumbProvider::DxtDecodeAlphaMap(BYTE* data, UINT dataLoc, BYTE alphaValues[16]) {
    BYTE alphaMap[8];
    alphaMap[0] = data[dataLoc + 0];
    alphaMap[1] = data[dataLoc + 1];

    if (alphaMap[0] > alphaMap[1]) {
        for (UINT i = 1; i < 7; i++) {
            alphaMap[i + 1] = (BYTE)(((ULONG)(7 - i) * (ULONG)alphaMap[0] + (ULONG)i * (ULONG)alphaMap[1]) / 7.0);
        }
    }
    else {
        for (UINT i = 1; i < 5; i++) {
            alphaMap[i + 1] = (BYTE)(((5 - i) * (UINT)alphaMap[0] + i * (UINT)alphaMap[1]) / 5.0);
        }
        alphaMap[6] = 0;
        alphaMap[7] = (BYTE)0xff;
    }

    ULONGLONG alphaBits = 0;

    for (int i = 2; i < 8; i++) { // pack the rest of the data into a single long for easy access
        alphaBits |= ((ULONGLONG) data[i + dataLoc]) << ((i - 2) * 8);
    }

    for (int i = 0; i < 16; i++) {
        alphaValues[i] = alphaMap[alphaBits & 0b111];
        alphaBits >>= 3;
    }
}

VOID CPapaThumbProvider::DecodeTexture(BYTE* data, USHORT width, USHORT height, BYTE format, BYTE* dst) {

    int heightZero = height - 1;

    if (format == 1) { // RGBA8888
        for (UINT y = 0; y < height; y++) {
            for (UINT x = 0; x < width; x++) {
                UINT i = (x + (heightZero - y) * width) * 4;
                UINT i2 = (x + y * width) * 4;
                dst[i] = data[i2];
                dst[i + 1] = data[i2 + 1];
                dst[i + 2] = data[i2 + 2];
                dst[i + 3] = data[i2 + 3];
            }
        }
    }
    else if (format == 2) { // RGBX8888
        for (UINT y = 0; y < height; y++) {
            for (UINT x = 0; x < width; x++) {
                UINT i = (x + (heightZero - y) * width) * 4;
                UINT i2 = (x + y * width) * 4;
                dst[i] = data[i2];
                dst[i + 1] = data[i2 + 1];
                dst[i + 2] = data[i2 + 2];
                dst[i + 3] = 255;
            }
        }
    }
    else if (format == 3) { // BGRA8888
        for (UINT y = 0; y < height; y++) {
            for (UINT x = 0; x < width; x++) {
                UINT i = (x + (heightZero - y) * width) * 4;
                UINT i2 = (x + y * width) * 4;
                dst[i] = data[i2 + 2];
                dst[i + 1] = data[i2 + 1];
                dst[i + 2] = data[i2];
                dst[i + 3] = data[i2 + 3];
            }
        }
    }
    else if (format == 4) { // DXT1
        UINT bufferLoc = 0;
        BYTE colours[4][3];
        for (UINT y = 0; y < height; y += 4) {
            for (UINT x = 0; x < width; x += 4) {

                DxtDecodeColourMap(data, bufferLoc, colours);
                bufferLoc += 4;

                UINT bits = 0;
                bits |= data[bufferLoc++] << 0;
                bits |= data[bufferLoc++] << 8;
                bits |= data[bufferLoc++] << 16;
                bits |= data[bufferLoc++] << 24;

                for (UINT yy = 0; yy < 4; yy++) {
                    for (UINT xx = 0; xx < 4; xx++) { // copy our colour data into the array
                        UINT colourIndex = bits & 0b11;
                        if (yy + y < height && xx + y < width) {
                            UINT idx = (xx + x + (heightZero - (yy + y)) * width) * 4;
                            BYTE* col = colours[colourIndex];
                            dst[idx] = col[0];
                            dst[idx + 1] = col[1];
                            dst[idx + 2] = col[2];
                            dst[idx + 3] = 255;
                        }
                        bits >>= 2;
                    }
                }
            }
        }
    }
    else if (format == 6) { // DXT5
        UINT bufferLoc = 0;
        BYTE alphaValues[16];
        BYTE colours[4][3];
        for (UINT y = 0; y < height; y += 4) {
            for (UINT x = 0; x < width; x += 4) {

                DxtDecodeAlphaMap(data, bufferLoc, alphaValues);
                bufferLoc += 8;

                DxtDecodeColourMap(data, bufferLoc, colours);
                bufferLoc += 4;

                UINT bits = 0;
                bits |= data[bufferLoc++] << 0;
                bits |= data[bufferLoc++] << 8;
                bits |= data[bufferLoc++] << 16;
                bits |= data[bufferLoc++] << 24;

                for (UINT yy = 0; yy < 4; yy++) {
                    for (UINT xx = 0; xx < 4; xx++) { // copy our colour data into the array
                        UINT colourIndex = bits & 0b11;
                        if (yy + y < height && xx + x < width) {
                            UINT idx = (xx + x + (heightZero - (yy + y)) * width) * 4;
                            BYTE* col = colours[colourIndex];
                            dst[idx] = col[0];
                            dst[idx + 1] = col[1];
                            dst[idx + 2] = col[2];
                            dst[idx + 3] = alphaValues[xx + yy * 4];
                        }
                        bits >>= 2;
                    }
                }
            }
        }
    }
    else if (format == 13) {
        for (UINT y = 0; y < height; y++) {
            for (UINT x = 0; x < width; x++) {
                UINT idx = x + (heightZero - y) * width * 4;
                UINT idx2 = x + y * width;
                dst[idx] = data[idx2]; // R
                dst[idx + 1] = 0; // G
                dst[idx + 2] = 0; // B
                dst[idx + 3] = 0; // A
            }
        }
    }
    else {
        for (UINT y = 0; y < height; y++) {
            for (UINT x = 0; x < width; x++) {
                UINT idx = x + (heightZero - y) * width * 4;
                dst[idx] = 1; // R
                dst[idx + 1] = 0; // G
                dst[idx + 2] = 0; // B
                dst[idx + 3] = 255; // A
            }
        }
    }
}

// HRESULT CPapaThumbProvider::RescaleImage()

// IThumbnailProvider
IFACEMETHODIMP CPapaThumbProvider::GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pdwAlpha)
{

    if (cx < 0) { // won't buid unless i do something with cx....
        return E_INVALIDARG;
    }

    BYTE header[0x68];

    HRESULT result = S_OK;

    result = _pStream->Read(header, 0x68, NULL);

    if (result != S_OK) {
        return E_INVALIDARG;
    }

    BYTE papaMagic[] = "apaP";

    // is the header valid?
    if (memcmp(header, papaMagic, 4) != 0) {
        return E_INVALIDARG;
    }

    SHORT numTextures = *(((SHORT*)header) + 5);

    if (numTextures <= 0) {
        return E_INVALIDARG;
    }

    ULONGLONG textureOffset = *(((ULONGLONG*)header) + 5);

    LARGE_INTEGER seek = LARGE_INTEGER();
    seek.QuadPart = textureOffset;

    result = _pStream->Seek(seek, STREAM_SEEK_SET, NULL);

    if (result != S_OK) {
        return E_INVALIDARG;
    }

    BYTE textureHeader[24];

    result = _pStream->Read(textureHeader, 24, NULL);

    if (result != S_OK) {
        return E_INVALIDARG;
    }

    BYTE format = textureHeader[2];
    USHORT width = *(((USHORT*)textureHeader) + 2);
    USHORT height = *(((USHORT*)textureHeader) + 3);

    ULONGLONG dataSize = *(((ULONGLONG*)textureHeader) + 1);
    ULONGLONG dataOffset = *(((ULONGLONG*)textureHeader) + 2);

    LARGE_INTEGER seekTex = LARGE_INTEGER();
    seekTex.QuadPart = dataOffset;

    result = _pStream->Seek(seekTex, STREAM_SEEK_SET, NULL);

    if (result != S_OK) {
        return E_INVALIDARG;
    }

    BYTE* data = (BYTE*)malloc(dataSize);
    if (data == NULL) {
        return E_OUTOFMEMORY;
    }


    result = _pStream->Read(data, (ULONG)dataSize, NULL);

    if (result != S_OK) {
        free(data);
        return E_INVALIDARG;
    }

    
    BITMAPINFO bmi = { sizeof(bmi.bmiHeader) };
    bmi.bmiHeader.biWidth = width;
    bmi.bmiHeader.biHeight = height;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    BYTE* pBits;
    HBITMAP hbmp = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, reinterpret_cast<void**>(&pBits), NULL, 0);
    DecodeTexture(data, width, height, format, pBits);

    *phbmp = hbmp;
    *pdwAlpha = WTSAT_ARGB;

    // swap R and B
    for (UINT y = 0; y < height; y++) {
        for (UINT x = 0; x < width; x++) {
            UINT idx = (x + y * width) * 4;
            BYTE t = pBits[idx];
            pBits[idx] = pBits[idx + 2];
            pBits[idx + 2] = t;
        }
    }


    free(data);
    return S_OK;
}
