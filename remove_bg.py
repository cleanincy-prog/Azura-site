#!/usr/bin/env python3
"""
Убирает фон с PNG файлов используя u2net (полная модель, высокое качество).
Требует: pip install onnxruntime pillow numpy
"""
import urllib.request
import numpy as np
import onnxruntime as ort
from pathlib import Path
from PIL import Image

MODEL_URL  = "https://github.com/danielgatis/rembg/releases/download/v0.0.0/u2net.onnx"
MODEL_PATH = Path.home() / ".u2net" / "u2net.onnx"
FOLDER     = Path(__file__).parent

def download_model():
    MODEL_PATH.parent.mkdir(exist_ok=True)
    if MODEL_PATH.exists():
        print("Модель уже скачана ✓")
        return
    print("Скачиваю AI модель (~170 MB), подождите...")
    def progress(count, block, total):
        pct = count * block / total * 100
        print(f"\r  {min(pct,100):.1f}%", end="", flush=True)
    urllib.request.urlretrieve(MODEL_URL, MODEL_PATH, reporthook=progress)
    print("\nМодель скачана ✓\n")

def preprocess(img):
    img = img.convert("RGB").resize((320, 320))
    x = np.array(img, dtype=np.float32) / 255.0
    mean = np.array([0.485, 0.456, 0.406], dtype=np.float32)
    std  = np.array([0.229, 0.224, 0.225], dtype=np.float32)
    x = (x - mean) / std
    return x.transpose(2, 0, 1)[None]

def get_mask(session, img):
    w, h = img.size
    inp  = preprocess(img)
    pred = session.run(None, {"input.1": inp})[0][0][0]
    # Нормализуем и применяем чёткий порог
    pred = (pred - pred.min()) / (pred.max() - pred.min() + 1e-8)
    # Делаем края резче
    pred = np.clip((pred - 0.3) / 0.4, 0, 1)
    mask = Image.fromarray((pred * 255).astype(np.uint8)).resize((w, h), Image.LANCZOS)
    return mask

def process():
    download_model()
    session = ort.InferenceSession(str(MODEL_PATH))
    pngs = sorted(FOLDER.glob("*.png"))
    print(f"Обрабатываю {len(pngs)} файлов...\n")
    done = errors = 0

    for i, path in enumerate(pngs, 1):
        try:
            print(f"[{i}/{len(pngs)}] {path.name}", end=" ... ", flush=True)
            img  = Image.open(path).convert("RGBA")
            mask = get_mask(session, img)
            img.putalpha(mask)
            img.save(path, "PNG")
            done += 1
            print("✓")
        except Exception as e:
            errors += 1
            print(f"✗ {e}")

    print(f"\n✅ Готово: {done} файлов, ошибок: {errors}")

if __name__ == "__main__":
    process()
