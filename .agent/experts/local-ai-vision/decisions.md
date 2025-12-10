# Local AI Vision - Decisions

## Status: DEFERRED

This feature area is deferred to post-MVP. Research is preserved for future implementation.

## Decision Log

### 1. Defer AI Image Descriptions

**Decision:** Focus MVP on comment navigation; defer image AI
**Date:** December 2025
**Status:** Final

**Rationale:**
- Comment navigation provides immediate value
- AI vision requires significant additional complexity
- User's primary need is comment accessibility
- Can add AI features in future version

---

### 2. Target Hardware (When Implemented)

**Decision:** Optimize for Surface Laptop 4 with AMD Ryzen
**Date:** December 2025
**Status:** Preserved for future

**Rationale:**
- User's development machine
- AMD GPU can accelerate inference
- Florence-2 model identified as promising

**Research:**
- `research/local-vision-model-analysis.md`
- `research/surface_laptop4_complete_guide.md`

---

## Future Implementation Notes

When we revisit AI vision:

1. **Model:** Florence-2 (Microsoft, runs locally)
2. **Optimization:** DirectML for AMD GPU acceleration
3. **Integration:** Async processing to avoid blocking NVDA
4. **Features:**
   - Image description on demand
   - Chart/graph interpretation
   - Alt text generation suggestions

## Research Files

- `research/local-vision-model-analysis.md` - Model comparison and selection
- `research/surface_laptop4_complete_guide.md` - Hardware optimization guide
