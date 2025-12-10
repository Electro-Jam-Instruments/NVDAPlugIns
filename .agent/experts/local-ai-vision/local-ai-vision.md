# Local AI Vision - Expert Knowledge

## Status: DEFERRED - POST-MVP FEATURE

**This feature is NOT part of the MVP implementation.**

This document preserves research and architectural decisions for future implementation of AI-powered image descriptions. The MVP focuses on comment navigation only.

See `decisions.md` for the decision to defer this feature.

## Reference Files

**Related Documentation:**
- `decisions.md` - Decision log for deferring AI vision features
- `research/local-vision-model-analysis.md` - Model comparison and selection research
- `research/surface_laptop4_complete_guide.md` - Hardware optimization guide for AMD GPU

## Overview

The goal is to provide AI-powered image descriptions for PowerPoint slides, running locally (no cloud API) for privacy and offline use.

## Target Hardware

**Primary:** Surface Laptop 4 with AMD Ryzen 7 4980U
- 8 cores / 16 threads
- AMD Radeon Graphics (integrated)
- DirectML support for GPU acceleration

## Recommended Model

### Florence-2 (Microsoft)

**Why Florence-2:**
- Open source (MIT license)
- Runs locally
- Multiple capabilities: captioning, OCR, object detection
- Reasonable size for consumer hardware
- DirectML compatible

**Variants:**
| Model | Size | Use Case |
|-------|------|----------|
| Florence-2-base | ~230M params | Fast inference |
| Florence-2-large | ~770M params | Better accuracy |

### Alternative: BLIP-2

- Good for general image captioning
- Larger than Florence-2
- May be slower on integrated GPU

## Implementation Approach

### Architecture

```
PowerPoint Image → Extract → Resize → Model → Description → NVDA Speech
                     ↓
              Async Processing (don't block NVDA)
```

### Key Considerations

1. **Async processing** - Model inference takes time; don't freeze NVDA
2. **Caching** - Cache descriptions per image hash
3. **Queue management** - Handle multiple images gracefully
4. **Fallback** - If model fails, announce "Image, no description available"

### DirectML Optimization

```python
import onnxruntime as ort

# Use DirectML execution provider for AMD GPU
session = ort.InferenceSession(
    "model.onnx",
    providers=['DmlExecutionProvider', 'CPUExecutionProvider']
)
```

## Integration with NVDA

### Gesture Binding

```python
@script(
    description="Describe current image",
    gesture="kb:NVDA+shift+i",
    category="PowerPoint AI"
)
def script_describeImage(self, gesture):
    image = self._get_current_image()
    if image:
        self._queue_description(image)
        ui.message("Analyzing image...")
```

### Background Processing

```python
import threading
import queue

class ImageDescriber:
    def __init__(self):
        self._queue = queue.Queue()
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    def _worker(self):
        while True:
            image, callback = self._queue.get()
            description = self._model.describe(image)
            callback(description)

    def describe_async(self, image, callback):
        self._queue.put((image, callback))
```

## Performance Targets

| Metric | Target |
|--------|--------|
| First inference | < 5 seconds |
| Cached response | < 100ms |
| Memory usage | < 2GB |
| GPU utilization | Prefer GPU when available |

## Future Features

1. **Chart interpretation** - "Bar chart showing sales by quarter"
2. **Table description** - "Table with 5 rows and 3 columns"
3. **SmartArt** - "Organizational chart with 4 levels"
4. **Alt text suggestions** - Generate for content creators

## Next Steps (When Resuming This Deferred Feature)

1. Set up Florence-2 with ONNX Runtime + DirectML
2. Create async wrapper for NVDA integration
3. Implement image extraction from PowerPoint shapes
4. Add caching layer
5. Test with screen reader users
