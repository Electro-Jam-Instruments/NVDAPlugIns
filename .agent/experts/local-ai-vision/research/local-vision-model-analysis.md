# Local AI Vision Model Integration Analysis
## For NVDA PowerPoint Plugin Image Description

**Document Version:** 1.0
**Date:** December 3, 2025
**Purpose:** Strategic analysis and recommendation for integrating local vision-language models for accessibility image descriptions

---

## Executive Summary

This document analyzes two primary approaches for integrating local AI vision models into an NVDA PowerPoint plugin for generating image descriptions: **Ollama + LLaVA** and **Florence-2 + NPU/DirectML**. After comprehensive research, **Option A (Ollama + LLaVA) is recommended for MVP development** due to significantly simpler deployment, better accessibility for blind users during installation, and adequate performance for the use case.

**Key Recommendation:** Start with Ollama + LLaVA for MVP, with Moondream as a faster alternative for users with limited hardware. Consider Florence-2 optimization as a future enhancement once the core plugin is stable.

---

## Table of Contents

1. [Option A: Ollama + LLaVA Analysis](#option-a-ollama--llava-analysis)
2. [Option B: Florence-2 + NPU Analysis](#option-b-florence-2--npu-analysis)
3. [Option C: Alternative Models](#option-c-alternative-models)
4. [Performance Comparison Matrix](#performance-comparison-matrix)
5. [Deployment Complexity Analysis](#deployment-complexity-analysis)
6. [Recommendation and Justification](#recommendation-and-justification)
7. [Implementation Roadmap](#implementation-roadmap)
8. [Risk Assessment](#risk-assessment)
9. [References](#references)

---

## Option A: Ollama + LLaVA Analysis

### Overview

Ollama provides a local LLM server with REST API access, supporting vision models like LLaVA. The server runs on localhost:11434 and provides a simple HTTP interface for image analysis.

### Installation Requirements

| Requirement | Specification |
|-------------|---------------|
| Operating System | Windows 10 22H2+ or Windows 11 |
| Minimum RAM | 8 GB (16 GB recommended) |
| Minimum VRAM | 4-6 GB with quantization |
| Disk Space | ~4.7 GB for llava:7b model |
| Network | Required only for initial model download |

### Installation Process

**Standard Installation:**
```
1. Download OllamaSetup.exe from ollama.com
2. Run installer (requires admin privileges)
3. Open command prompt/terminal
4. Run: ollama pull llava:7b
5. Model downloads automatically (~4.7 GB)
```

**Silent Installation (for deployment):**
- Use `/SILENT` or `/VERYSILENT` flags with InnoSetup
- Alternative: Standalone ZIP file (ollama-windows-amd64.zip) for service-based deployment
- Note: Custom path configuration during silent install has reported issues

### Performance Expectations

| Metric | llava:7b | llava:13b |
|--------|----------|-----------|
| Model Size | ~4.7 GB | ~8 GB |
| First Image (cold) | 5-10 seconds | 10-15 seconds |
| Cached Images | <0.1 seconds | <0.1 seconds |
| VRAM Usage (quantized) | ~5-6 GB | ~8-10 GB |
| VRAM Usage (FP16) | ~14-16 GB | ~26+ GB |

**CPU-Only Performance:**
- Intel Core i7-1355U: ~7.5 tokens/second
- AMD Ryzen 5 4600G: ~12.3 tokens/second
- Significantly slower than GPU, but functional

### Accuracy for PowerPoint Content

LLaVA 1.6 includes specific improvements for document/presentation content:
- Trained on additional document, chart, and diagram datasets
- Supports up to 4x more pixels for detail recognition
- Improved text recognition and reasoning capabilities
- 85.1% relative score compared to GPT-4 on multimodal tasks
- 92.53% accuracy on Science QA (with GPT-4 synergy)

**Chart/Diagram Specific:**
- Markdown-style output suits complex VQA tasks
- Document and chart understanding explicitly supported
- Can extract insights from charts, diagrams, and text-heavy images

### API Integration

**Python Example:**
```python
import ollama
import base64

def describe_image(image_path: str) -> str:
    with open(image_path, 'rb') as f:
        image_data = base64.b64encode(f.read()).decode()

    response = ollama.chat(
        model='llava:7b',
        messages=[{
            'role': 'user',
            'content': 'Describe this PowerPoint slide image for a blind user. Focus on text content, charts, diagrams, and visual elements.',
            'images': [image_data]
        }]
    )
    return response['message']['content']
```

**Error Handling:**
- `RequestError`: Connection/request issues
- `ResponseError`: HTTP errors with status codes
- `ConnectionError`: Server unavailable
- Automatic retry and fallback supported via LiteLLM

### Accessibility Concerns

**CRITICAL FINDING:** Ollama's Windows GUI application has documented accessibility issues with screen readers (NVDA and others):
- Toggle buttons don't expose on/off state
- Unlabeled buttons exist
- Radio button selections not announced

**MITIGATION:** The REST API and command-line interface are fully accessible. The plugin should:
1. Communicate only via REST API (localhost:11434)
2. Provide installation instructions in accessible text format
3. Use command-line installation where possible

### Offline Operation

- Fully offline after initial model download
- No data leaves the user's device
- Local-first architecture ideal for privacy-sensitive applications
- Models stored locally in Ollama's data directory

---

## Option B: Florence-2 + NPU Analysis

### Overview

Florence-2 is Microsoft's compact vision-language model with 230M (base) or 770M (large) parameters. Can be deployed via ONNX Runtime with DirectML (AMD GPU) or QNN (Qualcomm NPU) execution providers.

### Technical Specifications

| Specification | Florence-2-base | Florence-2-large |
|---------------|-----------------|------------------|
| Parameters | 230 million | 770 million |
| Model Size (FP32) | ~900 MB | ~3 GB |
| Model Size (INT8) | ~230 MB | ~770 MB |
| Model Size (INT4) | ~115 MB | ~385 MB |
| Memory at Inference | ~2 GB | ~4 GB |

### Quantization and Deployment

**CRITICAL LIMITATION DISCOVERED:**
- **DirectML does NOT support INT8 quantization**
- DirectML only supports INT4 (q4) quantized models
- For INT8, must use CUDA or TensorRT execution providers

**Available ONNX Models:**
- onnx-community/Florence-2-base (HuggingFace)
- onnx-community/Florence-2-base-ft (fine-tuned)
- Includes: Vision Encoder, Embed Tokens, Encoder, Decoder, Decoder Merged

**Quantization Options for DirectML:**
```
FP32 (full precision) - Supported
FP16 (half precision) - Supported
INT4 (4-bit quantization) - Supported
INT8 (8-bit quantization) - NOT SUPPORTED on DirectML
```

### NPU/QNN Execution Provider

**Status:** Experimental/Limited Support

**Requirements:**
- Windows 11 24H2 or 25H2
- Qualcomm Snapdragon X Elite/Plus processor
- QNN SDK and drivers
- ARM64 Python environment

**Current Limitations:**
- Quantization utilities only supported on x86_64
- No documented Florence-2 + QNN examples found
- Would require custom ONNX conversion and optimization
- Limited community support for this specific combination

**Recent Update (KB5072095):**
- Microsoft updated QNN execution provider to version 1.8.21.0
- Improved hardware-accelerated AI on Snapdragon platforms

### Performance Expectations

| Configuration | Latency Estimate |
|---------------|------------------|
| AMD GPU (DirectML, FP16) | 200-400 ms |
| AMD GPU (DirectML, INT4) | 150-300 ms |
| Snapdragon NPU (QNN) | 50-100 ms (theoretical) |
| CPU Only | 1-3 seconds |

**Note:** The claimed 50-100ms NPU latency is theoretical. No benchmarks for Florence-2 specifically on QNN were found.

### Accuracy Considerations

- 5% worse than 8B Idefics2 generalist model (impressive for size)
- Strong for captioning, OCR, and object detection
- Less tested specifically on chart/diagram understanding
- May require prompt engineering for accessibility descriptions

### Deployment Complexity

**High Complexity Factors:**
1. ONNX model conversion required
2. Multiple ONNX files to manage (5 components)
3. DirectML/QNN execution provider setup
4. Python environment with specific dependencies
5. No simple installer like Ollama
6. Platform-specific optimization needed

---

## Option C: Alternative Models

### Moondream2 (Recommended Alternative)

| Specification | Value |
|---------------|-------|
| Parameters | 1.86 billion |
| Model Size | 3.7 GB |
| VRAM Required | 4.3 GB (4 GB quantized) |
| Disk Space | 5 GB |

**Advantages:**
- Runs on CPU/MPS "quite fast"
- ONNX Runtime Web support (browser deployment possible)
- Successful deployments on Raspberry Pi
- Privacy-focused (local inference)

**Moondream 3.0 (September 2025):**
- MoE architecture: 9B total, 2B activated
- 32k context window
- Significant accuracy improvements:
  - COCO detection: 51.2 (+20.7 points)
  - ScreenSpot UI localization: 80.4 F1 (+20.1)
  - DocVQA: 79.3, TextVQA: 76.3, OCRBench: 61.2

**Ollama Support:**
```bash
ollama pull moondream
```

### Qwen2.5-VL

| Variant | VRAM Required | Performance |
|---------|---------------|-------------|
| 3B | ~6 GB | Good for lightweight use |
| 7B | ~8+ GB | Balanced performance |
| 72B | 48+ GB | Not suitable for consumer hardware |

**Advantages:**
- Available on Ollama: `ollama pull qwen2.5vl`
- Quantized versions achieve >99% accuracy at 8-bit
- Up to 3.5x speedup with quantization
- Strong document understanding

**Performance:**
- 3B model: GPU utilization ~6% on laptop with 8GB VRAM
- 7B model: 2.83 req/s, 2581 tokens/s output throughput

### BLIP-2 / InstructBLIP

**Status:** Not Recommended for This Use Case

**Issues:**
- No official ONNX support
- BLIP-2 ONNX conversion "not supported" by HuggingFace Optimum
- More complex deployment than alternatives
- Larger models without proportional accuracy gains for this use case

---

## Performance Comparison Matrix

| Model | Size | VRAM | First Image | Cached | Chart Accuracy | Installation |
|-------|------|------|-------------|--------|----------------|--------------|
| LLaVA 7b (Ollama) | 4.7 GB | 5-6 GB | 5-10s | <0.1s | Good | Simple |
| LLaVA 13b (Ollama) | 8 GB | 8-10 GB | 10-15s | <0.1s | Better | Simple |
| Florence-2-base (ONNX) | ~230 MB | ~2 GB | 200-400ms | Similar | Moderate | Complex |
| Moondream2 (Ollama) | 3.7 GB | 4.3 GB | 3-5s | <0.1s | Good | Simple |
| Qwen2.5-VL 3B (Ollama) | ~3 GB | ~6 GB | 4-8s | <0.1s | Good | Simple |

**Latency Validation:**
- Ollama estimates (5-10s first, <0.1s cached): **REALISTIC** based on benchmarks
- Florence-2 DirectML (150-300ms): **REALISTIC** for FP16/INT4, but setup is complex
- Florence-2 NPU (50-100ms): **UNVALIDATED** - no benchmarks found for this specific config

---

## Deployment Complexity Analysis

### For Blind Users with NVDA

**Option A (Ollama) - MODERATE COMPLEXITY:**
1. Download installer from ollama.com (accessible website)
2. Run OllamaSetup.exe (standard Windows installer)
3. Open accessible terminal (Windows Terminal or CMD)
4. Type: `ollama pull llava:7b`
5. Wait for download (~4.7 GB)
6. Plugin handles API communication

**Accessibility Notes:**
- Installer is standard Windows installer (generally accessible)
- Terminal commands are fully accessible with NVDA
- GUI app has accessibility issues (not needed for plugin)
- Progress reporting during download may not be accessible

**Option B (Florence-2 + ONNX) - HIGH COMPLEXITY:**
1. Install Python 3.8+ with pip
2. Install onnxruntime-directml package
3. Download 5 ONNX model files from HuggingFace
4. Configure DirectML execution provider
5. Handle model loading in Python
6. Manage multiple dependencies

**Accessibility Notes:**
- Python installation requires careful navigation
- HuggingFace download interface less accessible
- Multiple configuration steps prone to errors
- No single installer solution

### Dependency Management

| Approach | Dependencies | Update Burden |
|----------|--------------|---------------|
| Ollama + LLaVA | Ollama only | Low (auto-updates available) |
| Florence-2 + ONNX | Python, onnxruntime, numpy, PIL, etc. | High |
| Moondream (Ollama) | Ollama only | Low |

### Fallback Strategies

**When AI is Unavailable:**
1. Check if Ollama server is running (API ping)
2. Provide fallback to alt-text if present
3. Indicate "AI description unavailable"
4. Log error for troubleshooting
5. Offer to retry after user starts Ollama

---

## Recommendation and Justification

### Primary Recommendation: Ollama + LLaVA (Option A)

**For MVP Development:**
```
Model: llava:7b (or moondream for lower-spec hardware)
Server: Ollama localhost:11434
Integration: REST API via Python requests
Fallback: Alt-text extraction from PowerPoint
```

### Justification

1. **Accessibility-First Design**
   - Command-line installation accessible with screen readers
   - No GUI interaction required for core functionality
   - Clear text-based feedback during setup

2. **Simplest Deployment**
   - Single installer + single command
   - No Python environment management needed by end users
   - Automatic GPU detection and optimization

3. **Adequate Performance**
   - 5-10s first image acceptable for accessibility use case
   - Sub-second cached responses for repeated analysis
   - Works on consumer hardware (8GB+ RAM, 6GB+ VRAM ideal)

4. **Strong Accuracy for Use Case**
   - LLaVA 1.6 specifically trained on documents/charts
   - Improved text recognition capabilities
   - Proven performance on visual reasoning tasks

5. **Excellent Error Handling**
   - Well-documented Python library
   - Clear exception types for different failure modes
   - Timeout configuration for reliability

6. **Future Flexibility**
   - Can switch models with single command
   - Moondream available as lighter alternative
   - Qwen2.5-VL available for accuracy improvements

### Secondary Recommendation: Moondream as Lightweight Alternative

For users with limited hardware (8GB RAM, integrated GPU):
```bash
ollama pull moondream
```
- Smaller footprint (3.7 GB vs 4.7 GB)
- Lower VRAM requirement (4.3 GB vs 5-6 GB)
- Acceptable accuracy for accessibility descriptions

---

## Implementation Roadmap

### Phase 1: MVP (Recommended)

**Timeline:** 2-4 weeks

1. Implement Ollama REST API client in Python/Add-on
2. Add server availability check (ping localhost:11434)
3. Implement image extraction from PowerPoint
4. Create accessible image description prompts
5. Integrate descriptions into NVDA speech output
6. Provide fallback to alt-text when Ollama unavailable
7. Create accessible installation guide

**Deliverables:**
- Working prototype with LLaVA integration
- Installation documentation for blind users
- Error handling and fallback mechanisms

### Phase 2: Optimization (Optional)

**Timeline:** 4-8 weeks after MVP

1. Add caching layer for repeated images
2. Implement background processing for multi-image slides
3. Add user preferences for model selection
4. Consider Florence-2 integration for faster responses
5. Performance profiling and optimization

### Phase 3: Advanced Features (Future)

**Timeline:** Post-MVP based on user feedback

1. Custom prompt engineering for specific content types
2. Integration with Moondream 3.0 for improved accuracy
3. NPU acceleration research (if hardware adoption increases)
4. Batch processing for entire presentations

---

## Risk Assessment

### Option A (Ollama + LLaVA) Risks

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Ollama installation fails | Low | High | Provide troubleshooting guide, manual install option |
| Model download interrupted | Medium | Low | Support resume, provide offline model installation |
| GPU not detected | Medium | Medium | Document CPU fallback, performance expectations |
| API changes in Ollama | Low | Medium | Pin to stable version, monitor updates |
| Insufficient hardware | Medium | High | Recommend Moondream, document minimum specs |

### Option B (Florence-2 + ONNX) Risks

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| DirectML compatibility issues | High | High | Extensive testing required |
| Complex installation fails | High | High | Detailed guides, installer script |
| INT8 not supported (confirmed) | Certain | Medium | Use INT4 instead |
| NPU support unavailable | High | Medium | Fall back to DirectML GPU |
| Model conversion issues | Medium | High | Use pre-converted community models |

### Overall Risk Comparison

```
Option A Risk Score: MODERATE (manageable with documentation)
Option B Risk Score: HIGH (technical complexity, accessibility barriers)
```

---

## References

### Ollama and LLaVA

- [Ollama Official Documentation](https://docs.ollama.com/)
- [Ollama Windows Documentation](https://docs.ollama.com/windows)
- [LLaVA Official Page](https://llava-vl.github.io/)
- [Ollama LLaVA Model Page](https://ollama.com/library/llava:7b)
- [Ollama Python Library](https://github.com/ollama/ollama-python)
- [Ollama Accessibility Issues (GitHub)](https://github.com/ollama/ollama/issues/11806)
- [Ollama Hardware Guide](https://www.arsturn.com/blog/ollama-hardware-guide-what-you-need-to-run-llms-locally)
- [Best Ollama Models 2025](https://collabnix.com/best-ollama-models-in-2025-complete-performance-comparison/)
- [LLaVA GitHub Repository](https://github.com/haotian-liu/LLaVA)
- [Chart and Diagram Analysis with Ollama LLaVA](https://markaicode.com/ollama-llava-chart-diagram-analysis-guide/)

### Florence-2 and ONNX

- [ONNX Runtime DirectML Documentation](https://onnxruntime.ai/docs/execution-providers/DirectML-ExecutionProvider.html)
- [ONNX Runtime QNN Execution Provider](https://onnxruntime.ai/docs/execution-providers/QNN-ExecutionProvider.html)
- [Florence-2 ONNX Community Models](https://huggingface.co/onnx-community/Florence-2-base)
- [ONNX Runtime Quantization](https://onnxruntime.ai/docs/performance/model-optimizations/quantization.html)
- [Florence-2 Vision-Language Model](https://blog.roboflow.com/florence-2/)
- [Copilot+ PCs Developer Guide](https://learn.microsoft.com/en-us/windows/ai/npu-devices/)

### Alternative Models

- [Moondream Official](https://moondream.ai/)
- [Moondream GitHub](https://github.com/vikhyat/moondream)
- [Moondream HuggingFace](https://huggingface.co/vikhyatk/moondream2)
- [Qwen2.5-VL Ollama](https://ollama.com/library/qwen2.5vl)
- [Qwen2.5-VL Technical Report](https://arxiv.org/abs/2502.13923)
- [Vision Language Models 2025 Overview](https://huggingface.co/blog/vlms-2025)

### Performance and Benchmarks

- [LLM Hardware Benchmarking](https://medium.com/@ttio2tech_28094/local-large-language-models-hardware-benchmarking-ollama-benchmarks-cpu-gpu-macbooks-c696abbec613)
- [Ollama GPU Benchmarks RTX 4090](https://www.databasemart.com/blog/ollama-gpu-benchmark-rtx4090)
- [Qwen2.5-VL Benchmarks](https://debuggercafe.com/qwen2-5-vl/)
- [FastVLM Research](https://machinelearning.apple.com/research/fast-vision-language-models)

### Accessibility

- [NVDA Download](https://www.nvaccess.org/download/)
- [NVDA Features and Requirements](https://dss.sonoma.edu/nvda-features-and-system-requirements)

---

## Appendix A: Sample API Integration Code

### Ollama Python Client

```python
"""
Ollama LLaVA integration for NVDA PowerPoint Plugin
"""
import base64
import requests
from pathlib import Path
from typing import Optional

OLLAMA_API = "http://localhost:11434/api/generate"
DEFAULT_MODEL = "llava:7b"
TIMEOUT_SECONDS = 60

def check_ollama_available() -> bool:
    """Check if Ollama server is running."""
    try:
        response = requests.get("http://localhost:11434/api/tags", timeout=5)
        return response.status_code == 200
    except requests.ConnectionError:
        return False

def describe_image(
    image_path: str,
    model: str = DEFAULT_MODEL,
    prompt: Optional[str] = None
) -> str:
    """
    Generate an accessibility description for an image.

    Args:
        image_path: Path to the image file
        model: Ollama model to use
        prompt: Custom prompt (default optimized for accessibility)

    Returns:
        Text description of the image

    Raises:
        ConnectionError: If Ollama server unavailable
        ValueError: If image cannot be processed
    """
    if not check_ollama_available():
        raise ConnectionError("Ollama server not running. Please start Ollama.")

    image_data = Path(image_path).read_bytes()
    encoded_image = base64.b64encode(image_data).decode('utf-8')

    if prompt is None:
        prompt = """Describe this PowerPoint slide image for a blind user.
        Include:
        1. Any text content (headings, bullet points, labels)
        2. Charts or graphs (type, data trends, key values)
        3. Diagrams (structure, relationships, flow)
        4. Images or icons (what they depict)
        5. Layout and visual organization
        Be concise but comprehensive."""

    payload = {
        "model": model,
        "prompt": prompt,
        "images": [encoded_image],
        "stream": False
    }

    try:
        response = requests.post(OLLAMA_API, json=payload, timeout=TIMEOUT_SECONDS)
        response.raise_for_status()
        return response.json()["response"]
    except requests.Timeout:
        raise TimeoutError("Image analysis timed out. Try a smaller model.")
    except requests.RequestException as e:
        raise ValueError(f"Failed to process image: {e}")
```

### Error Handling Example

```python
def get_image_description_safe(image_path: str, alt_text: str = "") -> str:
    """
    Get image description with fallback to alt-text.

    Args:
        image_path: Path to image
        alt_text: Fallback alt-text from PowerPoint

    Returns:
        Description string (AI-generated or alt-text fallback)
    """
    try:
        return describe_image(image_path)
    except ConnectionError:
        if alt_text:
            return f"AI unavailable. Alt-text: {alt_text}"
        return "Image description unavailable. Ollama not running."
    except TimeoutError:
        if alt_text:
            return f"AI timed out. Alt-text: {alt_text}"
        return "Image analysis timed out."
    except ValueError as e:
        if alt_text:
            return f"Processing error. Alt-text: {alt_text}"
        return f"Could not process image: {e}"
```

---

## Appendix B: Installation Guide Template (Accessible)

```
OLLAMA INSTALLATION GUIDE FOR NVDA USERS
=========================================

This guide helps you install Ollama and LLaVA for AI image descriptions.

REQUIREMENTS:
- Windows 10 version 22H2 or newer
- 8 GB RAM minimum (16 GB recommended)
- 6 GB disk space
- Internet connection for initial download

STEP 1: DOWNLOAD OLLAMA
-----------------------
1. Press Windows key, type "Edge" or "Chrome", press Enter
2. Navigate to: ollama.com
3. Find the Download button (usually near top of page)
4. Download the Windows installer
5. The file is named OllamaSetup.exe

STEP 2: INSTALL OLLAMA
----------------------
1. Press Windows key + E to open File Explorer
2. Navigate to Downloads folder
3. Find OllamaSetup.exe and press Enter
4. Follow the installer prompts
5. Accept default options

STEP 3: DOWNLOAD THE AI MODEL
-----------------------------
1. Press Windows key, type "cmd", press Enter
2. In the command prompt, type:
   ollama pull llava:7b
3. Press Enter
4. Wait for download (about 4.7 GB)
5. You will see "success" when complete

STEP 4: VERIFY INSTALLATION
---------------------------
1. In command prompt, type:
   ollama list
2. Press Enter
3. You should hear "llava:7b" in the list

TROUBLESHOOTING:
- If "ollama" command not found, restart your computer
- If download fails, check internet connection and retry
- For slower computers, try: ollama pull moondream

The NVDA PowerPoint plugin will automatically connect to Ollama.
```

---

**Document End**
