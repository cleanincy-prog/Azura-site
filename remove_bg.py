#!/usr/bin/env python3
"""
Убирает фон с PNG файлов используя u2net + onnxruntime.
Не требует rembg, pymatting, numba или llvmlite.
"""
import urllib.request
import numpy as np
import onnxruntime as ort
from pathlib import Path
from PIL import Image

MODEL_URL = "https://github.com/danielgatis/rembg/releases/download/v0.0.0/u2netp.onnx"
MODEL_PATH = Path.home() / ".u2net" / "u2netp.onnx"
FOLDER = Path(__file__).parent

def download_model():
    MODEL_PATH.parent.mkdir(exist_ok=True)
    if MODEL_PATH.exists():
        print("Модель уже скачана ✓")
        return
    print("Скачиваю AI модель (~4 MB)...")
    urllib.request.urlretrieve(MODEL_URL, MODEL_PATH)
    print("Модель скачана ✓\n")

def preprocess(img):
    img = img.convert("RGB").resize((320, 320))
    x = np.array(img, dtype=np.float32) / 255.0
    mean = np.array([0.485, 0.456, 0.406], dtype=np.float32)
    std  = np.array([0.229, 0.224, 0.225], dtype=np.float32)
    x = (x - mean) / std
    return x.transpose(2, 0, 1)[None]

def predict(session, img):
    w, h = img.size
    inp = preprocess(img)
    out = session.run(None, {"input.1": inp})[0][0][0]
    mask = (out - out.min()) / (out.max() - out.min() + 1e-8)
    mask = Image.fromarray((mask * 255).astype(np.uint8)).resize((w, h), Image.LANCZOS)
    return mask

def remove_bg(img, mask):
    rgba = img.convert("RGBA")
    rgba.putalpha(mask)
    return rgba

def process():
    download_model()
    session = ort.InferenceSession(str(MODEL_PATH))
    pngs = sorted(FOLDER.glob("*.png"))
    print(f"Обрабатываю {len(pngs)} файлов...\n")
    done = errors = 0
    for i, path in enumerate(pngs, 1):
        try:
            print(f"[{i}/{len(pngs)}] {path.name}", end=" ... ", flush=True)
            img = Image.open(path)
            mask = predict(session, img)
            result = remove_bg(img, mask)
            result.save(path)
            done += 1
            print("✓")
        except Exception as e:
            errors += 1
            print(f"✗ {e}")
    print(f"\n✅ Готово: {done} файлов, ошибок: {errors}")

if __name__ == "__main__":
    process()
