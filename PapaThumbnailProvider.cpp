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
#include <Windows.h>
#include "ImgPapafile.c"

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
    HRESULT RescaleImageBilinear(HBITMAP*, HBITMAP*);
    FLOAT Lerp(FLOAT s, FLOAT e, FLOAT t);
    FLOAT Blerp(FLOAT c00, FLOAT c10, FLOAT c01, FLOAT c11, FLOAT tx, FLOAT ty);
    VOID Blit(HBITMAP*, HBITMAP*, LONG, LONG);
    VOID SwapBR(HBITMAP*);
    VOID SwapTopBottom(HBITMAP*);

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
        colours[2][0] = (BYTE)((2 * (ULONG64)colours[0][0] + (ULONG64)colours[1][0]) / 3.0);
        colours[2][1] = (BYTE)((2 * (ULONG64)colours[0][1] + (ULONG64)colours[1][1]) / 3.0);
        colours[2][2] = (BYTE)((2 * (ULONG64)colours[0][2] + (ULONG64)colours[1][2]) / 3.0);

        colours[3][0] = (BYTE)(((ULONG64)colours[0][0] + 2 * (ULONG64)colours[1][0]) / 3.0);
        colours[3][1] = (BYTE)(((ULONG64)colours[0][1] + 2 * (ULONG64)colours[1][1]) / 3.0);
        colours[3][2] = (BYTE)(((ULONG64)colours[0][2] + 2 * (ULONG64)colours[1][2]) / 3.0);
    }
    else {
        colours[2][0] = (BYTE)(((ULONG64)colours[0][0] + (ULONG64)colours[1][0]) / 2.0);
        colours[2][1] = (BYTE)(((ULONG64)colours[0][1] + (ULONG64)colours[1][1]) / 2.0);
        colours[2][2] = (BYTE)(((ULONG64)colours[0][2] + (ULONG64)colours[1][2]) / 2.0);

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
            alphaMap[i + 1] = (BYTE)(((ULONG64)(7 - i) * (ULONG64)alphaMap[0] + (ULONG64)i * (ULONG64)alphaMap[1]) / 7.0);
        }
    }
    else {
        for (UINT i = 1; i < 5; i++) {
            alphaMap[i + 1] = (BYTE)(((5 - i) * (ULONG64)alphaMap[0] + i * (ULONG64)alphaMap[1]) / 5.0);
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

// source:
// https://rosettacode.org/wiki/Bilinear_interpolation#C

FLOAT CPapaThumbProvider::Lerp(FLOAT s, FLOAT e, FLOAT t) {
    return s + (e - s) * t;
}

FLOAT CPapaThumbProvider::Blerp(FLOAT c00, FLOAT c10, FLOAT c01, FLOAT c11, FLOAT tx, FLOAT ty) {
    return Lerp(Lerp(c00, c10, tx), Lerp(c01, c11, tx), ty);
}

HRESULT CPapaThumbProvider::RescaleImageBilinear(HBITMAP* src, HBITMAP* dst) {

    DIBSECTION srcDib;
    GetObject(*src, sizeof(srcDib), (LPVOID)&srcDib);
    BYTE* srcPixels = (BYTE*)srcDib.dsBm.bmBits;
    LONG srcWidth = srcDib.dsBmih.biWidth;
    LONG srcHeight = srcDib.dsBmih.biHeight;

    DIBSECTION dstDib;
    GetObject(*dst, sizeof(dstDib), (LPVOID)&dstDib);
    BYTE* dstPixels = (BYTE*)dstDib.dsBm.bmBits;
    LONG dstWidth = dstDib.dsBmih.biWidth;
    LONG dstHeight = dstDib.dsBmih.biHeight;

    ULONG32* srcPixelsInt = (ULONG32*)srcPixels;
    ULONG32* dstPixelsInt = (ULONG32*)dstPixels;


    for (LONG y = 0; y < dstHeight; y++) {
        for (LONG x = 0; x < dstWidth; x++) {
            FLOAT gx = x / (FLOAT)(dstWidth) * (srcWidth - 0.5f);
            FLOAT gy = y / (FLOAT)(dstHeight) * (srcHeight - 0.5f);
            LONG gxi = (LONG)gx;
            LONG gyi = (LONG)gy;

            ULONG32 result = 0;
            ULONG32 c00 = srcPixelsInt[gxi + gyi * srcWidth];
            ULONG32 c10 = srcPixelsInt[gxi + 1 + gyi * srcWidth];
            ULONG32 c01 = srcPixelsInt[gxi + (gyi + 1) * srcWidth];
            ULONG32 c11 = srcPixelsInt[gxi + 1 + (gyi + 1) * srcWidth];
            for (LONG i = 0; i < 4; i++) {
                result |= (UCHAR)Blerp(   (FLOAT)((c00 >> (8 * i)) & 0xFF), (FLOAT)((c10 >> (8 * i)) & 0xFF), 
                                            (FLOAT)((c01 >> (8 * i)) & 0xFF), (FLOAT)((c11 >> (8 * i)) & 0xFF), 
                                            (FLOAT)(gx - gxi), (FLOAT)(gy - gyi)) << (8 * i);
            }
            dstPixelsInt[x + y * dstWidth] = result;
        }
    }
    return S_OK;
}

VOID CPapaThumbProvider::Blit(HBITMAP* src, HBITMAP* dst, LONG dx, LONG dy) {
    DIBSECTION srcDib;
    GetObject(*src, sizeof(srcDib), (LPVOID)&srcDib);
    BYTE* srcPixels = (BYTE*)srcDib.dsBm.bmBits;
    LONG srcWidth = srcDib.dsBmih.biWidth;
    LONG srcHeight = srcDib.dsBmih.biHeight;

    DIBSECTION dstDib;
    GetObject(*dst, sizeof(dstDib), (LPVOID)&dstDib);
    BYTE* dstPixels = (BYTE*)dstDib.dsBm.bmBits;
    LONG dstWidth = dstDib.dsBmih.biWidth;
    LONG dstHeight = dstDib.dsBmih.biHeight;

    // clamp the input to always be valid
    dx = min(max(dx, 0), dstWidth - srcWidth);
    dy = min(max(dy, 0), dstHeight - srcHeight);
    
    for (LONG y = dy; y < dy + srcHeight; y++) {
        for (LONG x = dx; x < dx + srcWidth; x++) {
            LONG sx = x - dx;
            LONG sy = y - dy;
            FLOAT srcAlpha = (FLOAT)(srcPixels[(sx + sy * srcWidth) * 4 + 3]) / 255.0f;
            FLOAT dstAlpha = (FLOAT)(dstPixels[(x + y * dstWidth) * 4 + 3]) / 255.0f;

            BYTE sred = srcPixels[(sx + sy * srcWidth) * 4];
            BYTE sgreen = srcPixels[(sx + sy * srcWidth) * 4 + 1];
            BYTE sblue = srcPixels[(sx + sy * srcWidth) * 4 + 2];

            BYTE dred = dstPixels[(x + y * dstWidth) * 4];
            BYTE dgreen = dstPixels[(x + y * dstWidth) * 4 + 1];
            BYTE dblue = dstPixels[(x + y * dstWidth) * 4 + 2];

            BYTE red = (BYTE)(sred * srcAlpha + dred * (1.0f - srcAlpha));
            BYTE green = (BYTE)(sgreen * srcAlpha + dgreen * (1.0f - srcAlpha));
            BYTE blue = (BYTE)(sblue * srcAlpha + dblue * (1.0f - srcAlpha));

            BYTE alpha = (BYTE) ((srcAlpha + (dstAlpha * (1.0f - srcAlpha))) * 255.0f);

            dstPixels[(x + y * dstWidth) * 4] = red;
            dstPixels[(x + y * dstWidth) * 4 + 1] = green;
            dstPixels[(x + y * dstWidth) * 4 + 2] = blue;
            dstPixels[(x + y * dstWidth) * 4 + 3] = alpha;
        }
    }   
}

VOID CPapaThumbProvider::SwapBR(HBITMAP* src) {
    DIBSECTION dib;
    GetObject(*src, sizeof(dib), (LPVOID)&dib);
    BYTE* pixels = (BYTE*)dib.dsBm.bmBits;
    LONG width = dib.dsBmih.biWidth;
    LONG height = dib.dsBmih.biHeight;

    // swap R and B
    for (LONG y = 0; y < height; y++) {
        for (LONG x = 0; x < width; x++) {
            LONG idx = (x + y * width) * 4;
            BYTE t = pixels[idx];
            pixels[idx] = pixels[idx + 2];
            pixels[idx + 2] = t;
        }
    }
}

void CPapaThumbProvider::SwapTopBottom(HBITMAP* src) {
    DIBSECTION dib;
    GetObject(*src, sizeof(dib), (LPVOID)&dib);
    BYTE* pixels = (BYTE*)dib.dsBm.bmBits;
    LONG width = dib.dsBmih.biWidth;
    LONG height = dib.dsBmih.biHeight;

    ULONG32* pixelsInt = (ULONG32*)pixels;

    for (LONG y = 0; y < height / 2; y++) {
        for (LONG x = 0; x < width; x++) {
            LONG idx = (x + y * width);
            LONG idx2 = (x + (height - y - 1) * width);
            ULONG32 t = pixelsInt[idx];
            pixelsInt[idx] = pixelsInt[idx2];
            pixelsInt[idx2] = t;
        }
    }
}

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

    BYTE* data = (BYTE*)malloc((size_t)dataSize);
    if (data == NULL) {
        return E_OUTOFMEMORY;
    }


    result = _pStream->Read(data, (ULONG)dataSize, NULL);

    if (result != S_OK) {
        free(data);
        return E_INVALIDARG;
    }

    // BGRA
    
    BITMAPINFO decompData = { sizeof(decompData.bmiHeader) };
    decompData.bmiHeader.biWidth = width;
    decompData.bmiHeader.biHeight = height;
    decompData.bmiHeader.biPlanes = 1;
    decompData.bmiHeader.biBitCount = 32;
    decompData.bmiHeader.biCompression = BI_RGB;

    BYTE* decompTexture;
    HBITMAP decompBitmap = CreateDIBSection(NULL, &decompData, DIB_RGB_COLORS, reinterpret_cast<void**>(&decompTexture), NULL, 0);
    DecodeTexture(data, width, height, format, decompTexture);

    // swap R and B
    SwapBR(&decompBitmap);

    // scale to desired size
    USHORT smaller = min(width, height);
    FLOAT factor = (FLOAT)cx / (FLOAT)smaller;

    BITMAPINFO bmi = { sizeof(bmi.bmiHeader) };
    bmi.bmiHeader.biWidth = (LONG)roundf(width * factor);
    bmi.bmiHeader.biHeight = (LONG)roundf(height * factor);
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    BYTE* pBits;
    HBITMAP scaledBitmap = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, reinterpret_cast<void**>(&pBits), NULL, 0);

    RescaleImageBilinear(&decompBitmap, &scaledBitmap);
    DeleteObject(decompBitmap); // free the old bitmap


    // load the papafile file image
    BITMAPINFO papafile = { sizeof(papafile.bmiHeader) };
    papafile.bmiHeader.biWidth = img_papafile.width;
    papafile.bmiHeader.biHeight = img_papafile.height;
    papafile.bmiHeader.biPlanes = 1;
    papafile.bmiHeader.biBitCount = (WORD)(img_papafile.bytes_per_pixel * 8);
    papafile.bmiHeader.biCompression = BI_RGB;
    BYTE* papafileBits;
    HBITMAP papafileBitmap = CreateDIBSection(NULL, &papafile, DIB_RGB_COLORS, reinterpret_cast<void**>(&papafileBits), NULL, 0);

    memcpy(papafileBits, img_papafile.pixel_data, (ULONG64)img_papafile.width * (ULONG64)img_papafile.height * (ULONG64)img_papafile.bytes_per_pixel);
    SwapBR(&papafileBitmap);
    SwapTopBottom(&papafileBitmap);

    const UINT offset = 1;
    
    Blit(&papafileBitmap, &scaledBitmap, bmi.bmiHeader.biWidth - papafile.bmiHeader.biWidth - offset, offset);

    *phbmp = scaledBitmap;
    *pdwAlpha = WTSAT_ARGB;

    free(data);
    return S_OK;
}
