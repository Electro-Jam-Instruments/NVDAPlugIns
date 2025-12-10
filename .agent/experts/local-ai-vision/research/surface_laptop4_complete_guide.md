# Florence-2 Model Optimization on Surface Laptop 4 (AMD)
## Complete Setup and Deployment Guide

---

## System Specifications

**Target System:** Surface Laptop 4  
**Processor:** AMD Ryzen 7 Microsoft Surface Edition (8 cores, 16 threads)  
**RAM:** 16 GB  
**Architecture:** x64  
**GPU:** AMD Integrated Graphics  
**OS:** Windows 11  

**Optimization Goal:** Create INT8 quantized Florence-2 model optimized for PowerPoint slides  
**Deployment Target:** Can be used on this AMD system OR transferred to ARM64 Snapdragon device  
**Total Time:** 2-4 hours (mostly automated)  

---

## Table of Contents

1. [Phase 0: Prerequisites Check](#phase-0-prerequisites-check)
2. [Phase 1: Environment Setup](#phase-1-environment-setup)
3. [Phase 2: Project Structure Setup](#phase-2-project-structure-setup)
4. [Phase 3: Download Base Model](#phase-3-download-base-model)
5. [Phase 4: Create Calibration Dataset](#phase-4-create-calibration-dataset)
6. [Phase 5: Run Base Model Quantization](#phase-5-run-base-model-quantization)
7. [Phase 6: Verification and Testing](#phase-6-verification-and-testing)
8. [Phase 7: Package for Deployment](#phase-7-package-for-deployment)
9. [Phase 8: Overnight Automation](#phase-8-overnight-automation)
10. [Appendix: Troubleshooting](#appendix-troubleshooting)

---

## Phase 0: Prerequisites Check

### Check System Requirements

```powershell
# check_system.ps1
Write-Host "=== System Check ===" -ForegroundColor Cyan

# Check Windows version
$os = Get-CimInstance Win32_OperatingSystem
Write-Host "`nOS: $($os.Caption) $($os.Version)"

# Check architecture
$arch = $env:PROCESSOR_ARCHITECTURE
Write-Host "Architecture: $arch"
if ($arch -ne "AMD64") {
    Write-Host "WARNING: Expected AMD64 architecture" -ForegroundColor Yellow
}

# Check RAM
$ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
Write-Host "RAM: $ram GB"
if ($ram -lt 15) {
    Write-Host "WARNING: Recommended 16GB+ RAM" -ForegroundColor Yellow
}

# Check free disk space
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
Write-Host "Free Disk Space (C:): $freeGB GB"
if ($freeGB -lt 20) {
    Write-Host "WARNING: Need at least 20GB free space" -ForegroundColor Red
} else {
    Write-Host "Disk space: OK" -ForegroundColor Green
}

# Check Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "`nPython: $pythonVersion"
    
    # Check if x64
    $pythonArch = python -c "import platform; print(platform.machine())" 2>&1
    Write-Host "Python Architecture: $pythonArch"
    
    if ($pythonArch -notmatch "AMD64|x86_64") {
        Write-Host "ERROR: Need x64 Python" -ForegroundColor Red
    } else {
        Write-Host "Python: OK" -ForegroundColor Green
    }
} catch {
    Write-Host "ERROR: Python not found. Please install Python 3.10 or 3.11" -ForegroundColor Red
    Write-Host "Download from: https://www.python.org/downloads/"
}

Write-Host "`n=== System Check Complete ===" -ForegroundColor Cyan
```

### Install Python (if needed)

**If Python is not installed:**

1. Download Python 3.11 (x64): https://www.python.org/downloads/windows/
2. During installation:
   - ✅ Check "Add Python to PATH"
   - ✅ Check "Install for all users"
   - Choose "Customize installation"
   - ✅ Check "pip"
   - ✅ Check "Add Python to environment variables"
3. Restart PowerShell after installation

**Verify Python:**
```powershell
python --version
python -c "import platform; print(platform.machine())"
# Should output: AMD64
```

---

## Phase 1: Environment Setup

### Create Project Directory

```powershell
# Create project root
New-Item -Path "C:\Florence2Optimization" -ItemType Directory -Force
Set-Location "C:\Florence2Optimization"

Write-Host "Project directory created: C:\Florence2Optimization" -ForegroundColor Green
```

### Create Virtual Environment

```powershell
# setup_environment.ps1

Write-Host "Creating virtual environment..." -ForegroundColor Cyan

# Create virtual environment
python -m venv olive_env

# Activate
.\olive_env\Scripts\Activate.ps1

# Upgrade pip
python -m pip install --upgrade pip

Write-Host "Virtual environment created and activated" -ForegroundColor Green
Write-Host "Location: $PWD\olive_env" -ForegroundColor Gray
```

### Install Dependencies

```powershell
# install_dependencies.ps1

Write-Host "Installing dependencies..." -ForegroundColor Cyan
Write-Host "This will take 5-10 minutes..." -ForegroundColor Yellow

# Create requirements file
$requirements = @"
# Core dependencies
olive-ai[cpu]==0.6.1
onnx==1.16.0
onnxruntime-directml==1.18.0
transformers==4.40.0
torch==2.2.0
torchvision==0.17.0

# Data handling
pillow==10.3.0
numpy==1.26.4
datasets==2.19.0

# Utilities
tqdm==4.66.2
psutil==5.9.8
huggingface-hub==0.23.0
"@

Set-Content -Path "requirements.txt" -Value $requirements

# Install packages
pip install -r requirements.txt

# Verify installations
Write-Host "`nVerifying installations..." -ForegroundColor Cyan

python -c "import olive; print(f'Olive: {olive.__version__}')"
python -c "import onnxruntime as ort; print(f'ONNX Runtime: {ort.__version__}')"
python -c "import transformers; print(f'Transformers: {transformers.__version__}')"
python -c "import torch; print(f'PyTorch: {torch.__version__}')"

# Check DirectML provider
python -c "import onnxruntime as ort; providers = ort.get_available_providers(); print(f'DirectML available: {\"DmlExecutionProvider\" in providers}')"

Write-Host "`nDependencies installed successfully!" -ForegroundColor Green
```

### Verify Installation

```python
# verify_installation.py

import sys
import onnxruntime as ort
import olive
from transformers import AutoProcessor
import torch
import platform

print("="*70)
print("INSTALLATION VERIFICATION")
print("="*70)

checks = []

# Python version
py_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
checks.append(("Python Version", py_version, "3.10+" in py_version or "3.11" in py_version))

# Architecture
arch = platform.machine()
checks.append(("Architecture", arch, arch in ["AMD64", "x86_64"]))

# Olive
checks.append(("Olive", olive.__version__, True))

# ONNX Runtime
checks.append(("ONNX Runtime", ort.__version__, True))

# DirectML provider
providers = ort.get_available_providers()
has_dml = "DmlExecutionProvider" in providers
checks.append(("DirectML Provider", str(has_dml), has_dml))

# Transformers
import transformers
checks.append(("Transformers", transformers.__version__, True))

# PyTorch
checks.append(("PyTorch", torch.__version__, True))

# Print results
print("\nComponent Status:")
all_passed = True
for name, value, passed in checks:
    status = "✓ PASS" if passed else "✗ FAIL"
    color = "\033[92m" if passed else "\033[91m"
    reset = "\033[0m"
    print(f"{color}{status}{reset} | {name}: {value}")
    if not passed:
        all_passed = False

print("\n" + "="*70)
if all_passed:
    print("✓ All checks passed! Ready to proceed.")
else:
    print("✗ Some checks failed. Please review errors above.")
print("="*70)

sys.exit(0 if all_passed else 1)
```

**Run verification:**
```powershell
python verify_installation.py
```

---

## Phase 2: Project Structure Setup

### Create Directory Structure

```powershell
# create_structure.ps1

Write-Host "Creating project structure..." -ForegroundColor Cyan

$directories = @(
    "models\florence2_base",
    "models\quantized\baseline",
    "models\quantized\powerpoint",
    "calibration_data\baseline\coco",
    "calibration_data\baseline\imagenet",
    "calibration_data\baseline\documents",
    "calibration_data\powerpoint\slides",
    "calibration_data\powerpoint\curated",
    "test_data\powerpoint\charts",
    "test_data\powerpoint\diagrams",
    "test_data\powerpoint\text",
    "test_data\powerpoint\mixed",
    "test_data\general",
    "scripts",
    "logs",
    "deployment_package",
    "temp"
)

foreach ($dir in $directories) {
    New-Item -Path $dir -ItemType Directory -Force | Out-Null
    Write-Host "Created: $dir" -ForegroundColor Gray
}

Write-Host "`nProject structure created!" -ForegroundColor Green

# Create .gitignore
$gitignore = @"
# Virtual environment
olive_env/

# Model files
*.onnx
*.onnx.data

# Data
calibration_data/
test_data/

# Temp files
temp/
*.tmp

# Logs
logs/
*.log

# Python cache
__pycache__/
*.pyc
*.pyo

# OS files
.DS_Store
Thumbs.db
"@

Set-Content -Path ".gitignore" -Value $gitignore

Write-Host "Created .gitignore" -ForegroundColor Gray
```

**Run structure creation:**
```powershell
.\create_structure.ps1
```

---

## Phase 3: Download Base Model

### Download Florence-2 ONNX Model

```python
# scripts/download_florence2.py

import os
import sys
from pathlib import Path
from huggingface_hub import snapshot_download, hf_hub_download
from transformers import AutoProcessor
import torch

print("="*70)
print("DOWNLOADING FLORENCE-2 BASE MODEL")
print("="*70)

model_id = "microsoft/Florence-2-base"
output_dir = Path("models/florence2_base")
output_dir.mkdir(parents=True, exist_ok=True)

print(f"\nModel ID: {model_id}")
print(f"Output directory: {output_dir}")

# Step 1: Download processor/tokenizer
print("\n[1/3] Downloading processor and tokenizer...")
try:
    processor = AutoProcessor.from_pretrained(
        model_id,
        trust_remote_code=True
    )
    processor.save_pretrained(output_dir)
    print("✓ Processor downloaded")
except Exception as e:
    print(f"✗ Error downloading processor: {e}")
    sys.exit(1)

# Step 2: Download PyTorch model
print("\n[2/3] Downloading PyTorch model...")
try:
    from transformers import AutoModelForCausalLM
    
    model = AutoModelForCausalLM.from_pretrained(
        model_id,
        trust_remote_code=True,
        torch_dtype=torch.float32
    )
    print("✓ Model downloaded")
except Exception as e:
    print(f"✗ Error downloading model: {e}")
    sys.exit(1)

# Step 3: Convert to ONNX
print("\n[3/3] Converting to ONNX format...")
try:
    import torch.onnx
    
    # Prepare dummy inputs
    dummy_pixel_values = torch.randn(1, 3, 224, 224)
    dummy_input_ids = torch.randint(0, 1000, (1, 10))
    
    # Export
    onnx_path = output_dir / "florence2_base.onnx"
    
    model.eval()
    with torch.no_grad():
        torch.onnx.export(
            model,
            (dummy_pixel_values, dummy_input_ids),
            str(onnx_path),
            input_names=["pixel_values", "input_ids"],
            output_names=["output"],
            dynamic_axes={
                "pixel_values": {0: "batch"},
                "input_ids": {0: "batch", 1: "sequence"}
            },
            opset_version=17,
            do_constant_folding=True
        )
    
    print(f"✓ ONNX model saved to: {onnx_path}")
    
    # Check file size
    size_mb = onnx_path.stat().st_size / (1024 * 1024)
    print(f"  Model size: {size_mb:.2f} MB")
    
except Exception as e:
    print(f"✗ Error converting to ONNX: {e}")
    sys.exit(1)

print("\n" + "="*70)
print("✓ Florence-2 base model ready!")
print("="*70)
```

**Run download:**
```powershell
python scripts\download_florence2.py
```

**Expected output:**
- Model file: `models/florence2_base/florence2_base.onnx` (~900MB)
- Processor files: `models/florence2_base/*.json`
- Total time: 15-30 minutes

---

## Phase 4: Create Calibration Dataset

### Download Calibration Images

```python
# scripts/create_calibration_dataset.py

import os
import sys
from pathlib import Path
from datasets import load_dataset
from PIL import Image
import json
from tqdm import tqdm

print("="*70)
print("CREATING CALIBRATION DATASET")
print("="*70)

output_dir = Path("calibration_data/baseline")
output_dir.mkdir(parents=True, exist_ok=True)

all_images = []

# 1. COCO validation set (200 images)
print("\n[1/3] Downloading COCO images...")
coco_dir = output_dir / "coco"
coco_dir.mkdir(exist_ok=True)

try:
    dataset = load_dataset("detection-datasets/coco", split="validation", streaming=True)
    
    count = 0
    for idx, sample in enumerate(dataset):
        if count >= 200:
            break
        
        try:
            img = sample['image']
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            img_path = coco_dir / f"coco_{idx:04d}.jpg"
            img.save(img_path, quality=95)
            all_images.append(str(img_path))
            count += 1
            
            if count % 50 == 0:
                print(f"  Downloaded {count}/200 images")
        except Exception as e:
            print(f"  Warning: Skipped image {idx}: {e}")
            continue
    
    print(f"✓ Downloaded {count} COCO images")
    
except Exception as e:
    print(f"✗ Error downloading COCO: {e}")
    print("Continuing with other datasets...")

# 2. ImageNet samples (100 images)
print("\n[2/3] Downloading ImageNet samples...")
imagenet_dir = output_dir / "imagenet"
imagenet_dir.mkdir(exist_ok=True)

try:
    dataset = load_dataset("ILSVRC/imagenet-1k", split="validation", streaming=True, trust_remote_code=True)
    
    count = 0
    for idx, sample in enumerate(dataset):
        if count >= 100:
            break
        
        try:
            img = sample['image']
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            img_path = imagenet_dir / f"imagenet_{idx:04d}.jpg"
            img.save(img_path, quality=95)
            all_images.append(str(img_path))
            count += 1
            
            if count % 25 == 0:
                print(f"  Downloaded {count}/100 images")
        except Exception as e:
            print(f"  Warning: Skipped image {idx}: {e}")
            continue
    
    print(f"✓ Downloaded {count} ImageNet images")
    
except Exception as e:
    print(f"✗ Error downloading ImageNet: {e}")
    print("Continuing with other datasets...")

# 3. Document/Chart images (100 images)
print("\n[3/3] Downloading document/chart images...")
docs_dir = output_dir / "documents"
docs_dir.mkdir(exist_ok=True)

try:
    # ChartQA dataset has charts and diagrams
    dataset = load_dataset("HuggingFaceM4/ChartQA", split="validation", streaming=True)
    
    count = 0
    for idx, sample in enumerate(dataset):
        if count >= 100:
            break
        
        try:
            img = sample['image']
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            img_path = docs_dir / f"doc_{idx:04d}.jpg"
            img.save(img_path, quality=95)
            all_images.append(str(img_path))
            count += 1
            
            if count % 25 == 0:
                print(f"  Downloaded {count}/100 images")
        except Exception as e:
            print(f"  Warning: Skipped image {idx}: {e}")
            continue
    
    print(f"✓ Downloaded {count} document images")
    
except Exception as e:
    print(f"✗ Error downloading documents: {e}")
    print("Note: Some datasets may require authentication")

# Save manifest
manifest = {
    "total_images": len(all_images),
    "sources": {
        "coco": len(list((output_dir / "coco").glob("*.jpg"))),
        "imagenet": len(list((output_dir / "imagenet").glob("*.jpg"))),
        "documents": len(list((output_dir / "documents").glob("*.jpg")))
    },
    "images": all_images
}

manifest_path = output_dir / "calibration_manifest.json"
with open(manifest_path, 'w') as f:
    json.dump(manifest, f, indent=2)

print("\n" + "="*70)
print(f"✓ Calibration dataset created!")
print(f"  Total images: {len(all_images)}")
print(f"  Manifest: {manifest_path}")
print("="*70)
```

**Run dataset creation:**
```powershell
python scripts\create_calibration_dataset.py
```

**Expected:**
- ~400 images downloaded
- Time: 20-40 minutes
- Storage: ~2GB

### Preprocess Calibration Data

```python
# scripts/preprocess_calibration.py

import os
import sys
from pathlib import Path
from PIL import Image
import json
from transformers import AutoProcessor
from tqdm import tqdm
import numpy as np

print("="*70)
print("PREPROCESSING CALIBRATION DATA")
print("="*70)

# Load manifest
manifest_path = Path("calibration_data/baseline/calibration_manifest.json")
with open(manifest_path) as f:
    manifest = json.load(f)

print(f"\nTotal images to process: {manifest['total_images']}")

# Load processor
print("Loading Florence-2 processor...")
processor = AutoProcessor.from_pretrained(
    "models/florence2_base",
    trust_remote_code=True
)

# Process each image
valid_images = []
processed_count = 0

print("\nProcessing images...")
for img_path in tqdm(manifest['images']):
    try:
        # Load and validate image
        img = Image.open(img_path)
        
        # Convert to RGB if needed
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Check dimensions
        if img.size[0] < 50 or img.size[1] < 50:
            print(f"Warning: Skipping small image {img_path}")
            continue
        
        # Test preprocessing
        inputs = processor(
            text="<MORE_DETAILED_CAPTION>",
            images=img,
            return_tensors="pt"
        )
        
        # Validate tensor shapes
        if inputs['pixel_values'].shape != (1, 3, 224, 224):
            print(f"Warning: Unexpected shape for {img_path}")
            continue
        
        valid_images.append(img_path)
        processed_count += 1
        
    except Exception as e:
        print(f"Error processing {img_path}: {e}")
        continue

# Update manifest
manifest['valid_images'] = valid_images
manifest['valid_count'] = len(valid_images)
manifest['processed_count'] = processed_count

with open(manifest_path, 'w') as f:
    json.dump(manifest, f, indent=2)

print("\n" + "="*70)
print(f"✓ Preprocessing complete!")
print(f"  Valid images: {len(valid_images)}/{manifest['total_images']}")
print(f"  Manifest updated: {manifest_path}")
print("="*70)
```

**Run preprocessing:**
```powershell
python scripts\preprocess_calibration.py
```

---

## Phase 5: Run Base Model Quantization

### Create Data Reader for Olive

```python
# scripts/calibration_data_reader.py

import numpy as np
from PIL import Image
from pathlib import Path
import json
from transformers import AutoProcessor

class CalibrationDataReader:
    """Data reader for ONNX Runtime quantization calibration"""
    
    def __init__(self, data_dir, batch_size=1):
        self.data_dir = Path(data_dir)
        self.batch_size = batch_size
        
        # Load manifest
        manifest_path = self.data_dir / "calibration_manifest.json"
        with open(manifest_path) as f:
            manifest = json.load(f)
        
        self.image_paths = manifest.get('valid_images', manifest['images'])
        self.current_idx = 0
        
        # Load processor
        self.processor = AutoProcessor.from_pretrained(
            "models/florence2_base",
            trust_remote_code=True
        )
        
        print(f"Calibration data reader initialized: {len(self.image_paths)} images")
    
    def get_next(self):
        """Get next batch of calibration data"""
        if self.current_idx >= len(self.image_paths):
            return None
        
        batch_pixel_values = []
        batch_input_ids = []
        
        for _ in range(self.batch_size):
            if self.current_idx >= len(self.image_paths):
                break
            
            # Load image
            img_path = self.image_paths[self.current_idx]
            img = Image.open(img_path).convert('RGB')
            
            # Process with Florence-2 processor
            inputs = self.processor(
                text="<MORE_DETAILED_CAPTION>",
                images=img,
                return_tensors="np"
            )
            
            batch_pixel_values.append(inputs['pixel_values'])
            batch_input_ids.append(inputs['input_ids'])
            
            self.current_idx += 1
        
        if not batch_pixel_values:
            return None
        
        # Stack into batch
        return {
            "pixel_values": np.concatenate(batch_pixel_values, axis=0),
            "input_ids": np.concatenate(batch_input_ids, axis=0)
        }
    
    def rewind(self):
        """Reset to beginning"""
        self.current_idx = 0
```

### Quantization Script

```python
# scripts/quantize_baseline.py

import os
import sys
from pathlib import Path
import onnxruntime as ort
from onnxruntime.quantization import quantize_static, QuantType, QuantFormat
import time
import json

# Import custom data reader
sys.path.append(str(Path(__file__).parent))
from calibration_data_reader import CalibrationDataReader

print("="*70)
print("BASELINE MODEL QUANTIZATION")
print("="*70)

# Paths
model_path = Path("models/florence2_base/florence2_base.onnx")
output_path = Path("models/quantized/baseline/florence2_baseline_int8.onnx")
output_path.parent.mkdir(parents=True, exist_ok=True)

print(f"\nInput model: {model_path}")
print(f"Output model: {output_path}")

# Verify input model exists
if not model_path.exists():
    print(f"✗ Error: Model not found at {model_path}")
    sys.exit(1)

# Create data reader
print("\nInitializing calibration data...")
data_reader = CalibrationDataReader(
    data_dir="calibration_data/baseline",
    batch_size=1
)

# Quantization settings
print("\nQuantization settings:")
print("  Mode: Static INT8")
print("  Per-channel: True")
print("  Activation type: UINT8")
print("  Weight type: INT8")
print("  Format: QDQ (Quantize-Dequantize)")

# Run quantization
print("\nStarting quantization...")
print("This will take 60-90 minutes on your AMD Ryzen 7...")
start_time = time.time()

try:
    quantize_static(
        model_input=str(model_path),
        model_output=str(output_path),
        calibration_data_reader=data_reader,
        quant_format=QuantFormat.QDQ,
        per_channel=True,
        weight_type=QuantType.QInt8,
        activation_type=QuantType.QUInt8,
        optimize_model=True,
        extra_options={
            'ActivationSymmetric': False,
            'WeightSymmetric': True,
            'CalibMovingAverage': True,
            'CalibMovingAverageConstant': 0.01,
        }
    )
    
    elapsed_time = time.time() - start_time
    
    print(f"\n✓ Quantization complete!")
    print(f"  Time: {elapsed_time/60:.1f} minutes")
    
    # Check model size
    original_size = model_path.stat().st_size / (1024 * 1024)
    quantized_size = output_path.stat().st_size / (1024 * 1024)
    reduction = (1 - quantized_size/original_size) * 100
    
    print(f"  Original size: {original_size:.2f} MB")
    print(f"  Quantized size: {quantized_size:.2f} MB")
    print(f"  Size reduction: {reduction:.1f}%")
    
    # Save metadata
    metadata = {
        "model_type": "florence2_baseline_quantized",
        "quantization": "INT8",
        "calibration_dataset": "baseline",
        "calibration_samples": len(data_reader.image_paths),
        "original_size_mb": round(original_size, 2),
        "quantized_size_mb": round(quantized_size, 2),
        "size_reduction_percent": round(reduction, 2),
        "quantization_time_minutes": round(elapsed_time/60, 2),
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    
    metadata_path = output_path.parent / "baseline_metadata.json"
    with open(metadata_path, 'w') as f:
        json.dump(metadata, f, indent=2)
    
    print(f"  Metadata saved: {metadata_path}")
    
except Exception as e:
    print(f"\n✗ Quantization failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Verify model
print("\nVerifying quantized model...")
try:
    session = ort.InferenceSession(str(output_path))
    
    print("Model inputs:")
    for input in session.get_inputs():
        print(f"  {input.name}: {input.shape} ({input.type})")
    
    print("Model outputs:")
    for output in session.get_outputs():
        print(f"  {output.name}: {output.shape} ({output.type})")
    
    print("\n✓ Model verification passed!")
    
except Exception as e:
    print(f"\n✗ Model verification failed: {e}")
    sys.exit(1)

print("\n" + "="*70)
print("✓ BASELINE QUANTIZATION COMPLETE")
print("="*70)
```

**Run quantization:**
```powershell
python scripts\quantize_baseline.py
```

**Expected:**
- Time: 60-90 minutes
- Output: `models/quantized/baseline/florence2_baseline_int8.onnx` (~230MB)
- Progress updates every few minutes

---

## Phase 6: Verification and Testing

### Quick Inference Test

```python
# scripts/test_baseline_model.py

import os
import sys
from pathlib import Path
import onnxruntime as ort
from transformers import AutoProcessor
from PIL import Image
import time
import numpy as np

print("="*70)
print("BASELINE MODEL TESTING")
print("="*70)

# Load quantized model
model_path = Path("models/quantized/baseline/florence2_baseline_int8.onnx")

print(f"\nLoading model: {model_path}")

# Setup DirectML provider for AMD GPU
providers = [
    'DmlExecutionProvider',  # AMD GPU acceleration
    'CPUExecutionProvider'   # Fallback
]

sess_options = ort.SessionOptions()
sess_options.graph_optimization_level = ort.GraphOptimizationLevel.ORT_ENABLE_ALL

session = ort.InferenceSession(
    str(model_path),
    sess_options=sess_options,
    providers=providers
)

active_providers = session.get_providers()
print(f"Active providers: {active_providers}")

if 'DmlExecutionProvider' in active_providers:
    print("✓ AMD GPU acceleration enabled (DirectML)")
else:
    print("⚠ Using CPU fallback")

# Load processor
processor = AutoProcessor.from_pretrained(
    "models/florence2_base",
    trust_remote_code=True
)

# Create test image
print("\nCreating test image...")
test_img = Image.new('RGB', (800, 600), color=(100, 150, 200))
from PIL import ImageDraw, ImageFont
draw = ImageDraw.Draw(test_img)
draw.text((50, 50), "Test Slide", fill='white')
test_img_path = Path("temp/test_slide.png")
test_img_path.parent.mkdir(exist_ok=True)
test_img.save(test_img_path)

# Run inference test
print("\nRunning inference test...")

# Warmup (5 iterations)
print("Warming up...")
img = Image.open(test_img_path).convert('RGB')
inputs = processor(
    text="<MORE_DETAILED_CAPTION>",
    images=img,
    return_tensors="np"
)

for _ in range(5):
    _ = session.run(None, {
        "pixel_values": inputs['pixel_values'],
        "input_ids": inputs['input_ids']
    })

# Benchmark (10 iterations)
print("Running benchmark (10 iterations)...")
latencies = []

for i in range(10):
    start = time.perf_counter()
    outputs = session.run(None, {
        "pixel_values": inputs['pixel_values'],
        "input_ids": inputs['input_ids']
    })
    latency = (time.perf_counter() - start) * 1000
    latencies.append(latency)
    print(f"  Iteration {i+1}: {latency:.2f}ms")

# Decode output
generated_ids = outputs[0]
caption = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

# Results
print("\n" + "="*70)
print("PERFORMANCE RESULTS")
print("="*70)
print(f"Mean Latency:   {np.mean(latencies):.2f} ms")
print(f"Median Latency: {np.median(latencies):.2f} ms")
print(f"Min Latency:    {np.min(latencies):.2f} ms")
print(f"Max Latency:    {np.max(latencies):.2f} ms")
print(f"Std Dev:        {np.std(latencies):.2f} ms")
print(f"\nGenerated caption: {caption}")
print("="*70)

# Performance assessment
mean_latency = np.mean(latencies)
if mean_latency < 200:
    print("\n✓ Performance: EXCELLENT (<200ms)")
elif mean_latency < 400:
    print("\n✓ Performance: GOOD (<400ms)")
else:
    print("\n⚠ Performance: ACCEPTABLE (>400ms)")
    print("  Note: This is expected on CPU. DirectML may improve performance.")

print("\n✓ Baseline model test complete!")
```

**Run test:**
```powershell
python scripts\test_baseline_model.py
```

**Expected Results:**
- AMD GPU (DirectML): 150-300ms per image
- CPU fallback: 300-600ms per image

---

## Phase 7: Package for Deployment

### Create Deployment Package

```python
# scripts/package_for_deployment.py

import os
import sys
from pathlib import Path
import shutil
import json
import time

print("="*70)
print("CREATING DEPLOYMENT PACKAGE")
print("="*70)

# Create package directory
pkg_dir = Path("deployment_package")
if pkg_dir.exists():
    print("Removing existing package...")
    shutil.rmtree(pkg_dir)

pkg_dir.mkdir()

print("\n[1/6] Copying quantized model...")
model_src = Path("models/quantized/baseline/florence2_baseline_int8.onnx")
model_dst = pkg_dir / "florence2_baseline_int8.onnx"
shutil.copy2(model_src, model_dst)
model_size = model_dst.stat().st_size / (1024 * 1024)
print(f"  Model size: {model_size:.2f} MB")

print("\n[2/6] Copying processor files...")
processor_dst = pkg_dir / "processor"
shutil.copytree("models/florence2_base", processor_dst, dirs_exist_ok=True)
# Remove ONNX file from processor folder
for onnx_file in processor_dst.glob("*.onnx"):
    onnx_file.unlink()

print("\n[3/6] Creating configuration files...")

# AMD deployment config
amd_config = {
    "model_file": "florence2_baseline_int8.onnx",
    "processor_path": "processor",
    "target_device": "amd_gpu",
    "execution_provider": "DmlExecutionProvider",
    "model_info": {
        "type": "florence2_vision",
        "quantization": "int8",
        "optimized_for": "general_images",
        "input_size": [224, 224]
    },
    "performance_targets": {
        "latency_ms": 300,
        "throughput_img_per_sec": 3
    }
}

with open(pkg_dir / "config_amd.json", 'w') as f:
    json.dump(amd_config, f, indent=2)

# ARM deployment config (if transferring to Snapdragon)
arm_config = {
    "model_file": "florence2_baseline_int8.onnx",
    "processor_path": "processor",
    "target_device": "snapdragon_npu",
    "execution_provider": "QNNExecutionProvider",
    "qnn_options": {
        "backend_path": "QnnHtp.dll",
        "htp_performance_mode": "burst",
        "enable_htp_weight_sharing": True
    },
    "model_info": {
        "type": "florence2_vision",
        "quantization": "int8",
        "optimized_for": "general_images",
        "input_size": [224, 224]
    },
    "performance_targets": {
        "latency_ms": 100,
        "throughput_img_per_sec": 10
    }
}

with open(pkg_dir / "config_arm.json", 'w') as f:
    json.dump(arm_config, f, indent=2)

print("\n[4/6] Creating inference service script...")

inference_script = '''
# inference_service.py
"""
Florence-2 Inference Service
Supports both AMD (DirectML) and ARM64 (QNN) deployments
"""

import onnxruntime as ort
from transformers import AutoProcessor
from PIL import Image
import json
import os
from pathlib import Path
import time

class Florence2InferenceService:
    """Universal inference service for Florence-2"""
    
    def __init__(self, config_path="config_amd.json"):
        # Load configuration
        with open(config_path) as f:
            self.config = json.load(f)
        
        print(f"Initializing Florence-2 Inference Service...")
        print(f"Target device: {self.config['target_device']}")
        
        # Setup providers based on target
        if self.config['target_device'] == 'amd_gpu':
            providers = [
                'DmlExecutionProvider',
                'CPUExecutionProvider'
            ]
        elif self.config['target_device'] == 'snapdragon_npu':
            qnn_options = self.config.get('qnn_options', {})
            providers = [
                ('QNNExecutionProvider', qnn_options),
                'CPUExecutionProvider'
            ]
        else:
            providers = ['CPUExecutionProvider']
        
        # Session options
        sess_options = ort.SessionOptions()
        sess_options.graph_optimization_level = ort.GraphOptimizationLevel.ORT_ENABLE_ALL
        
        # Load model
        self.session = ort.InferenceSession(
            self.config['model_file'],
            sess_options=sess_options,
            providers=providers
        )
        
        # Load processor
        self.processor = AutoProcessor.from_pretrained(
            self.config['processor_path'],
            trust_remote_code=True
        )
        
        # Verify active providers
        active = self.session.get_providers()
        print(f"Active providers: {active}")
        
        if self.config['execution_provider'] in active:
            print(f"✓ {self.config['execution_provider']} active")
        else:
            print(f"⚠ Falling back to CPU")
    
    def generate_caption(self, image_path, prompt="<MORE_DETAILED_CAPTION>"):
        """Generate caption for an image"""
        # Load image
        img = Image.open(image_path).convert("RGB")
        
        # Prepare input
        inputs = self.processor(
            text=prompt,
            images=img,
            return_tensors="np"
        )
        
        # Run inference
        start = time.perf_counter()
        outputs = self.session.run(None, {
            "pixel_values": inputs['pixel_values'],
            "input_ids": inputs['input_ids']
        })
        latency_ms = (time.perf_counter() - start) * 1000
        
        # Decode output
        caption = self.processor.batch_decode(
            outputs[0],
            skip_special_tokens=True
        )[0]
        
        return {
            "caption": caption,
            "latency_ms": latency_ms
        }
    
    def get_model_info(self):
        """Get model information"""
        return {
            "config": self.config,
            "providers": self.session.get_providers(),
            "input_names": [inp.name for inp in self.session.get_inputs()],
            "output_names": [out.name for out in self.session.get_outputs()]
        }

if __name__ == "__main__":
    import sys
    
    # Test inference
    service = Florence2InferenceService()
    
    # Model info
    info = service.get_model_info()
    print("\\nModel Information:")
    print(f"  Active providers: {info['providers']}")
    
    # Test with sample image if provided
    if len(sys.argv) > 1:
        test_image = sys.argv[1]
        result = service.generate_caption(test_image)
        print(f"\\nTest Results:")
        print(f"  Caption: {result['caption']}")
        print(f"  Latency: {result['latency_ms']:.2f}ms")
    else:
        print("\\nUsage: python inference_service.py <test_image_path>")
'''

with open(pkg_dir / "inference_service.py", 'w') as f:
    f.write(inference_script)

print("\n[5/6] Creating README...")

readme = '''# Florence-2 Baseline Model - Deployment Package

## Contents

- `florence2_baseline_int8.onnx` - Quantized INT8 model (~230MB)
- `processor/` - Tokenizer and processor configuration
- `config_amd.json` - Configuration for AMD systems (Surface Laptop 4)
- `config_arm.json` - Configuration for ARM64 systems (Snapdragon)
- `inference_service.py` - Universal inference service
- `test_deployment.py` - Deployment verification script

## Deployment Instructions

### On AMD System (Surface Laptop 4)

1. Install dependencies:
```bash
pip install onnxruntime-directml transformers pillow
```

2. Test deployment:
```bash
python test_deployment.py
```

3. Run inference:
```bash
python inference_service.py test_image.png
```

### On ARM64 System (Snapdragon)

1. Ensure QNN SDK is installed
2. Install dependencies:
```bash
pip install onnxruntime-qnn transformers pillow
```

3. Test deployment (using ARM config):
```bash
python test_deployment.py config_arm.json
```

## Performance Expectations

- AMD GPU (DirectML): 150-300ms per image
- Snapdragon NPU (QNN): 50-100ms per image
- CPU fallback: 300-600ms per image

## Model Information

- Type: Florence-2 Vision Language Model
- Quantization: INT8
- Input size: 224x224
- Calibration: General images (COCO, ImageNet, Documents)

## Integration

For NVDA plugin integration, use `inference_service.py` as follows:

```python
from inference_service import Florence2InferenceService

service = Florence2InferenceService("config_amd.json")
result = service.generate_caption("slide.png")
print(result["caption"])
```

## Support

For issues or questions, refer to the main documentation.
'''

with open(pkg_dir / "README.md", 'w') as f:
    f.write(readme)

print("\n[6/6] Creating deployment test script...")

test_script = '''
# test_deployment.py
"""Test deployment and verify model works correctly"""

import sys
from pathlib import Path
import json

print("="*70)
print("DEPLOYMENT VERIFICATION")
print("="*70)

# Check Python packages
print("\\n[1/4] Checking dependencies...")
try:
    import onnxruntime as ort
    print(f"  ✓ onnxruntime: {ort.__version__}")
    
    import transformers
    print(f"  ✓ transformers: {transformers.__version__}")
    
    from PIL import Image
    print(f"  ✓ pillow: OK")
    
except ImportError as e:
    print(f"  ✗ Missing dependency: {e}")
    print("  Install with: pip install onnxruntime-directml transformers pillow")
    sys.exit(1)

# Check files
print("\\n[2/4] Checking files...")
required_files = [
    "florence2_baseline_int8.onnx",
    "processor/config.json",
    "config_amd.json",
    "inference_service.py"
]

all_present = True
for file in required_files:
    exists = Path(file).exists()
    status = "✓" if exists else "✗"
    print(f"  {status} {file}")
    if not exists:
        all_present = False

if not all_present:
    print("\\n✗ Some files are missing!")
    sys.exit(1)

# Load model
print("\\n[3/4] Loading model...")
try:
    from inference_service import Florence2InferenceService
    
    config = "config_amd.json"
    if len(sys.argv) > 1:
        config = sys.argv[1]
    
    service = Florence2InferenceService(config)
    print("  ✓ Model loaded successfully")
    
except Exception as e:
    print(f"  ✗ Error loading model: {e}")
    sys.exit(1)

# Test inference
print("\\n[4/4] Testing inference...")
try:
    # Create test image
    test_img = Image.new('RGB', (400, 300), color='blue')
    test_path = Path("test_temp.png")
    test_img.save(test_path)
    
    # Run inference
    result = service.generate_caption(str(test_path))
    
    print(f"  ✓ Inference successful")
    print(f"    Latency: {result['latency_ms']:.2f}ms")
    print(f"    Caption: {result['caption'][:80]}...")
    
    # Cleanup
    test_path.unlink()
    
except Exception as e:
    print(f"  ✗ Inference failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\\n" + "="*70)
print("✓ ALL CHECKS PASSED - Deployment ready!")
print("="*70)
'''

with open(pkg_dir / "test_deployment.py", 'w') as f:
    f.write(test_script)

# Create metadata
metadata = {
    "package_version": "1.0.0",
    "model_type": "florence2_baseline_int8",
    "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    "source_system": "Surface Laptop 4 (AMD Ryzen 7)",
    "contents": {
        "model": "florence2_baseline_int8.onnx",
        "processor": "processor/",
        "configs": ["config_amd.json", "config_arm.json"],
        "scripts": ["inference_service.py", "test_deployment.py"]
    }
}

with open(pkg_dir / "package_metadata.json", 'w') as f:
    json.dump(metadata, f, indent=2)

# Create ZIP archive
print("\n[7/7] Creating ZIP archive...")
archive_name = "florence2_baseline_deployment"
shutil.make_archive(archive_name, 'zip', pkg_dir)

archive_path = Path(f"{archive_name}.zip")
archive_size = archive_path.stat().st_size / (1024 * 1024)

print("\n" + "="*70)
print("✓ DEPLOYMENT PACKAGE CREATED")
print("="*70)
print(f"Package: {archive_path}")
print(f"Size: {archive_size:.2f} MB")
print(f"\nContents:")
print(f"  - Quantized model (INT8)")
print(f"  - Processor configuration")
print(f"  - AMD and ARM configs")
print(f"  - Inference service")
print(f"  - Test scripts")
print("\nReady to deploy on AMD or ARM64 systems!")
print("="*70)
```

**Run packaging:**
```powershell
python scripts\package_for_deployment.py
```

**Output:**
- `deployment_package/` folder with all files
- `florence2_baseline_deployment.zip` (~280MB)

---

## Phase 8: Overnight Automation

### Master Automation Script

```powershell
# run_overnight_optimization.ps1

Write-Host "="*70 -ForegroundColor Cyan
Write-Host "OVERNIGHT OPTIMIZATION - FLORENCE-2 MODEL" -ForegroundColor Cyan
Write-Host "="*70 -ForegroundColor Cyan
Write-Host ""
Write-Host "This script will run for approximately 2-4 hours"
Write-Host "Progress will be logged to overnight_log.txt"
Write-Host ""

# Setup
$logFile = "logs\overnight_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$startTime = Get-Date

# Create logs directory
New-Item -Path "logs" -ItemType Directory -Force | Out-Null

function Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Add-Content -Path $logFile -Value $logMessage
}

Log "=== OVERNIGHT OPTIMIZATION STARTED ==="
Log "Start time: $startTime"
Log "Log file: $logFile"

# Activate virtual environment
Log "`n[STEP 0/7] Activating virtual environment..."
try {
    & ".\olive_env\Scripts\Activate.ps1"
    Log "Virtual environment activated"
} catch {
    Log "ERROR: Failed to activate virtual environment"
    Log $_.Exception.Message
    exit 1
}

# Step 1: Download Florence-2
Log "`n[STEP 1/7] Downloading Florence-2 base model..."
Log "Expected time: 15-30 minutes"
try {
    python scripts\download_florence2.py 2>&1 | Tee-Object -FilePath $logFile -Append
    if ($LASTEXITCODE -ne 0) { throw "Download failed" }
    Log "✓ Download complete"
} catch {
    Log "ERROR: Florence-2 download failed"
    Log $_.Exception.Message
    exit 1
}

# Step 2: Create calibration dataset
Log "`n[STEP 2/7] Creating calibration dataset..."
Log "Expected time: 20-40 minutes"
try {
    python scripts\create_calibration_dataset.py 2>&1 | Tee-Object -FilePath $logFile -Append
    if ($LASTEXITCODE -ne 0) { throw "Dataset creation failed" }
    Log "✓ Dataset created"
} catch {
    Log "ERROR: Calibration dataset creation failed"
    Log $_.Exception.Message
    exit 1
}

# Step 3: Preprocess calibration data
Log "`n[STEP 3/7] Preprocessing calibration data..."
Log "Expected time: 10-15 minutes"
try {
    python scripts\preprocess_calibration.py 2>&1 | Tee-Object -FilePath $logFile -Append
    if ($LASTEXITCODE -ne 0) { throw "Preprocessing failed" }
    Log "✓ Preprocessing complete"
} catch {
    Log "ERROR: Preprocessing failed"
    Log $_.Exception.Message
    exit 1
}

# Step 4: Quantize baseline model
Log "`n[STEP 4/7] Quantizing baseline model..."
Log "Expected time: 60-90 minutes (LONGEST STEP)"
Log "This is the main optimization step - please be patient"
try {
    python scripts\quantize_baseline.py 2>&1 | Tee-Object -FilePath $logFile -Append
    if ($LASTEXITCODE -ne 0) { throw "Quantization failed" }
    Log "✓ Quantization complete"
} catch {
    Log "ERROR: Quantization failed"
    Log $_.Exception.Message
    exit 1
}

# Step 5: Test baseline model
Log "`n[STEP 5/7] Testing baseline model..."
Log "Expected time: 5-10 minutes"
try {
    python scripts\test_baseline_model.py 2>&1 | Tee-Object -FilePath $logFile -Append
    if ($LASTEXITCODE -ne 0) { throw "Testing failed" }
    Log "✓ Testing complete"
} catch {
    Log "ERROR: Testing failed"
    Log $_.Exception.Message
    exit 1
}

# Step 6: Package for deployment
Log "`n[STEP 6/7] Creating deployment package..."
Log "Expected time: 2-5 minutes"
try {
    python scripts\package_for_deployment.py 2>&1 | Tee-Object -FilePath $logFile -Append
    if ($LASTEXITCODE -ne 0) { throw "Packaging failed" }
    Log "✓ Package created"
} catch {
    Log "ERROR: Packaging failed"
    Log $_.Exception.Message
    exit 1
}

# Step 7: Final summary
Log "`n[STEP 7/7] Generating summary..."

$endTime = Get-Date
$duration = $endTime - $startTime

Log "`n=== OVERNIGHT OPTIMIZATION COMPLETE ==="
Log "End time: $endTime"
Log "Total duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
Log ""
Log "DELIVERABLES:"
Log "  - Quantized model: models\quantized\baseline\florence2_baseline_int8.onnx"
Log "  - Deployment package: florence2_baseline_deployment.zip"
Log "  - Full log: $logFile"
Log ""
Log "NEXT STEPS:"
Log "  1. Review the log file for any warnings"
Log "  2. Test the deployment package"
Log "  3. Transfer to ARM device (if needed)"
Log ""
Log "✓ ALL STEPS COMPLETED SUCCESSFULLY"

# Beep to alert user (if present)
[console]::beep(1000, 300)
Start-Sleep -Milliseconds 200
[console]::beep(1200, 300)
Start-Sleep -Milliseconds 200
[console]::beep(1400, 500)

Write-Host ""
Write-Host "="*70 -ForegroundColor Green
Write-Host "OPTIMIZATION COMPLETE!" -ForegroundColor Green
Write-Host "="*70 -ForegroundColor Green
Write-Host ""
Write-Host "Check the log file for details: $logFile" -ForegroundColor Yellow
```

### Keep System Awake Script

```powershell
# keep_awake.ps1
# Run this in a SEPARATE PowerShell window before starting optimization

Write-Host "="*70 -ForegroundColor Cyan
Write-Host "SYSTEM AWAKE MANAGER" -ForegroundColor Cyan
Write-Host "="*70 -ForegroundColor Cyan
Write-Host ""
Write-Host "This script will keep your system awake during optimization"
Write-Host "Keep this window open until optimization completes"
Write-Host ""
Write-Host "Press Ctrl+C to stop and allow system to sleep again"
Write-Host ""

# Prevent system sleep
$code = @'
[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
'@

$ste = Add-Type -MemberDefinition $code -Name System -Namespace Win32 -PassThru

$ES_CONTINUOUS = [uint32]"0x80000000"
$ES_SYSTEM_REQUIRED = [uint32]"0x00000001"
$ES_AWAYMODE_REQUIRED = [uint32]"0x00000040"

# Keep system awake
$null = $ste::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_SYSTEM_REQUIRED)

Write-Host "✓ System will stay awake" -ForegroundColor Green
Write-Host ""
Write-Host "Monitoring... (Press Ctrl+C to exit)" -ForegroundColor Yellow

try {
    while ($true) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "`r[$timestamp] System awake - Optimization running..." -NoNewline
        Start-Sleep -Seconds 60
    }
} finally {
    # Reset on exit
    $null = $ste::SetThreadExecutionState($ES_CONTINUOUS)
    Write-Host "`n`nSystem sleep settings restored" -ForegroundColor Yellow
}
```

### Complete Overnight Workflow

**Before Bed (~10 PM):**

```powershell
# 1. Open PowerShell as Administrator
cd C:\Florence2Optimization

# 2. Start keep-awake script (in separate window)
# Open another PowerShell window:
.\keep_awake.ps1

# 3. Start optimization (in original window)
.\run_overnight_optimization.ps1

# 4. Go to sleep! 😴
```

**Next Morning (~7 AM):**

```powershell
# 1. Check if complete (look for beep/window)

# 2. Review log
cat logs\overnight_log_*.txt

# 3. Verify outputs exist
dir models\quantized\baseline\florence2_baseline_int8.onnx
dir florence2_baseline_deployment.zip

# 4. Stop keep-awake script (Ctrl+C in that window)

# 5. Test deployment
cd deployment_package
python test_deployment.py
```

---

## Appendix: Troubleshooting

### Common Issues and Solutions

#### Issue: Python Not Found
```powershell
# Solution: Install Python 3.11 x64
# Download from: https://www.python.org/downloads/
# Make sure to check "Add Python to PATH"
```

#### Issue: Out of Memory During Quantization
```powershell
# Solution 1: Close other applications
# Solution 2: Reduce calibration dataset size in create_calibration_dataset.py
# Change: count >= 200 → count >= 100 (for COCO section)
```

#### Issue: DirectML Not Working
```powershell
# Check GPU drivers
# Update AMD drivers from Windows Update or AMD website

# Verify DirectML installation
python -c "import onnxruntime as ort; print(ort.get_available_providers())"
# Should include 'DmlExecutionProvider'

# Reinstall if needed
pip uninstall onnxruntime-directml
pip install onnxruntime-directml
```

#### Issue: Download Timeouts
```powershell
# Increase timeout and retry
$env:HF_HUB_DOWNLOAD_TIMEOUT = "300"
python scripts\download_florence2.py
```

#### Issue: Disk Space
```powershell
# Check free space
Get-PSDrive C | Select-Object Free

# Clean up after optimization:
# - Delete calibration_data (saves ~2GB)
# - Keep only the final .zip file (saves ~1GB)
```

### Performance Optimization Tips

#### For Faster Quantization
1. Close all other applications
2. Disable Windows Defender real-time scanning (temporarily)
3. Use SSD for project directory
4. Ensure laptop is plugged in (performance mode)

#### For Better Inference Speed
1. Update AMD GPU drivers
2. Use performance power plan:
   ```powershell
   powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
   ```
3. Close background applications
4. Ensure DirectML provider is active

### Verification Checklist

After overnight optimization, verify:

- [ ] Log file shows "ALL STEPS COMPLETED SUCCESSFULLY"
- [ ] Model file exists: `models/quantized/baseline/florence2_baseline_int8.onnx`
- [ ] Model size is ~230MB
- [ ] Deployment zip exists: `florence2_baseline_deployment.zip`
- [ ] Zip size is ~280MB
- [ ] Test script passes: `python scripts\test_baseline_model.py`
- [ ] Latency is acceptable (<400ms)

### Contact and Support

- Review logs in `logs/` directory
- Check model metadata in `models/quantized/baseline/baseline_metadata.json`
- Refer to main documentation for advanced troubleshooting

---

## Summary: What You Get

### After Overnight Optimization

**Files Created:**
1. **Quantized Model** (~230MB)
   - `models/quantized/baseline/florence2_baseline_int8.onnx`
   - Ready for inference on AMD or ARM systems

2. **Deployment Package** (~280MB)
   - `florence2_baseline_deployment.zip`
   - Everything needed for production deployment
   - Works on Surface Laptop 4 (AMD) immediately
   - Can be transferred to ARM64 Snapdragon device

3. **Logs and Metadata**
   - Complete optimization log
   - Performance metrics
   - Model verification results

### Performance Expectations

| System | Provider | Expected Latency |
|--------|----------|------------------|
| Surface Laptop 4 (AMD) | DirectML | 150-300ms |
| Surface Laptop 4 (CPU) | CPU only | 300-600ms |
| Snapdragon (ARM64) | QNN NPU | 50-100ms |

### Ready for NVDA Integration

The deployment package includes everything needed to integrate with your NVDA PowerPoint plugin:

```python
# In your NVDA plugin:
from inference_service import Florence2InferenceService

service = Florence2InferenceService("config_amd.json")
result = service.generate_caption("powerpoint_slide.png")
# Returns: {"caption": "...", "latency_ms": 234.56}
```

---

## Quick Reference Commands

```powershell
# Setup (one-time)
cd C:\Florence2Optimization
python -m venv olive_env
.\olive_env\Scripts\Activate.ps1
pip install -r requirements.txt

# Overnight automation
.\keep_awake.ps1                    # Window 1
.\run_overnight_optimization.ps1    # Window 2

# Manual steps (if needed)
python scripts\download_florence2.py
python scripts\create_calibration_dataset.py
python scripts\preprocess_calibration.py
python scripts\quantize_baseline.py
python scripts\test_baseline_model.py
python scripts\package_for_deployment.py

# Testing
python scripts\test_baseline_model.py
cd deployment_package
python test_deployment.py

# Cleanup
deactivate
```

---

**Total Hands-On Time:** ~20 minutes  
**Total Compute Time:** ~2-4 hours (automated)  
**Storage Required:** ~10GB (can reduce to ~1GB after optimization)  
**Ready for:** AMD deployment immediately, ARM64 after transfer
