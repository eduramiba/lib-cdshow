#pragma once

#include <stdint.h>

namespace cds_format_selection {

enum ProcessingRank {
    PROCESSING_UNKNOWN = 0,
    PROCESSING_ARGB32 = 1,
    PROCESSING_MJPG = 2,
    PROCESSING_YUY2 = 3,
    PROCESSING_NV12 = 4,
    PROCESSING_RGB24 = 5,
    PROCESSING_RGB32 = 6
};

struct Score {
    uint32_t maxFps;
    int processingRank;
    uint32_t stableIndex;

    constexpr Score(uint32_t fps, int rank, uint32_t index)
        : maxFps(fps), processingRank(rank), stableIndex(index) {}
};

// Resolution is filtered by the caller. Prefer the highest advertised frame
// rate first; only use conversion/decode cost to break an FPS tie. The stable
// index makes selection deterministic when drivers publish duplicate modes.
constexpr bool IsBetter(const Score& candidate, const Score& current) {
    return candidate.maxFps != current.maxFps
        ? candidate.maxFps > current.maxFps
        : (candidate.processingRank != current.processingRank
            ? candidate.processingRank > current.processingRank
            : candidate.stableIndex < current.stableIndex);
}

static_assert(
    IsBetter(
        Score(30, PROCESSING_MJPG, 1),
        Score(5, PROCESSING_RGB32, 0)),
    "Frame rate must win over subtype processing cost");
static_assert(
    IsBetter(
        Score(30, PROCESSING_RGB32, 1),
        Score(30, PROCESSING_MJPG, 0)),
    "RGB32 must win an equal-FPS processing-cost tie");
static_assert(
    IsBetter(
        Score(30, PROCESSING_NV12, 1),
        Score(30, PROCESSING_YUY2, 0)),
    "NV12 must preserve the established equal-FPS YUV preference");
static_assert(
    IsBetter(
        Score(0, PROCESSING_RGB24, 2),
        Score(0, PROCESSING_MJPG, 1)),
    "Processing cost must rank modes when FPS is unknown");
static_assert(
    IsBetter(
        Score(30, PROCESSING_RGB32, 2),
        Score(30, PROCESSING_RGB32, 3)),
    "Stable format index must resolve identical scores");

} // namespace cds_format_selection
