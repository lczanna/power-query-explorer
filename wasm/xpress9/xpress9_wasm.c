/*
 * Minimal WASM wrapper for XPress9 decompression.
 * Exposes three functions to JavaScript:
 *   - xpress9_init()       -> creates a decoder context
 *   - xpress9_decompress() -> decompresses a single block
 *   - xpress9_free()       -> releases the decoder context
 *
 * Copyright (c) Microsoft Corporation (XPress9 library) — MIT License
 */
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include "xpress.h"
#include "xpress9.h"
#include <emscripten/emscripten.h>

#define UNREFERENCED_PARAMETER(P) (void)(P)

static void* my_alloc(void *ctx, int size) {
    UNREFERENCED_PARAMETER(ctx);
    return malloc(size);
}
static void my_free(void *ctx, void *addr) {
    UNREFERENCED_PARAMETER(ctx);
    free(addr);
}

static XPRESS9_DECODER g_decoder = NULL;

EMSCRIPTEN_KEEPALIVE
int xpress9_init(void) {
    if (g_decoder) return 1;
    XPRESS9_STATUS status = {0};
    g_decoder = Xpress9DecoderCreate(&status, NULL, my_alloc,
                                      XPRESS9_WINDOW_SIZE_LOG2_MAX, 0);
    if (!g_decoder || status.m_uStatus != Xpress9Status_OK) return 0;
    Xpress9DecoderStartSession(&status, g_decoder, 1);
    return (status.m_uStatus == Xpress9Status_OK) ? 1 : 0;
}

EMSCRIPTEN_KEEPALIVE
unsigned xpress9_decompress(const unsigned char *src, int srcLen,
                            unsigned char *dst, int dstLen) {
    if (!g_decoder) return 0;
    XPRESS9_STATUS status = {0};
    unsigned total = 0;

    Xpress9DecoderAttach(&status, g_decoder, src, srcLen);
    if (status.m_uStatus != Xpress9Status_OK) return 0;

    while (1) {
        unsigned written = 0, needed = 0;
        Xpress9DecoderFetchDecompressedData(&status, g_decoder,
                                            dst + total, dstLen - total,
                                            &written, &needed);
        if (status.m_uStatus != Xpress9Status_OK) { total = 0; break; }
        if (written == 0) break;
        total += written;
    }

    Xpress9DecoderDetach(&status, g_decoder, src, srcLen);
    return total;
}

EMSCRIPTEN_KEEPALIVE
void xpress9_free(void) {
    if (g_decoder) {
        XPRESS9_STATUS status = {0};
        Xpress9DecoderDestroy(&status, g_decoder, NULL, my_free);
        g_decoder = NULL;
    }
}
