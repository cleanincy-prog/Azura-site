#!/usr/bin/env python3
"""
Вырезает этикетки с фото бутылок (бар + вино).
Сохраняет label_FILENAME.png рядом с оригиналами.
Требует: pillow numpy (уже установлены)
"""
import re, numpy as np
from PIL import Image
from pathlib import Path

FOLDER = Path(__file__).parent
HTML   = FOLDER / "azura_menu_guide.html"

# ── 1. Вытащить имена файлов из секций bar и wine ──
def extract_imgs(html, section):
    # Данные хранятся в barItems / wineItems / kitchenItems / cocktailItems
    var_name = section + "Items"
    pattern = rf"const {var_name}\s*=\s*\[([\s\S]*?)\];\s*\n"
    m = re.search(pattern, html)
    if not m:
        print(f"  Переменная {var_name} не найдена")
        return []
    block = m.group(1)
    imgs = re.findall(r"img:['\"]([^'\"]+\.(png|jpg|jpeg))['\"]", block, re.IGNORECASE)
    return [f for f, _ in imgs if f]

# ── 2. Найти зону этикетки ──
def find_label_box(arr, crop_ratio=0.38):
    """Находит горизонтальную полосу с максимальной дисперсией (= этикетка)"""
    h, w = arr.shape[:2]
    gray = np.mean(arr[:, :, :3], axis=2).astype(np.float32)

    # Считаем дисперсию по полосам высотой 5% изображения
    strip_h = max(1, h // 20)
    variances = []
    for y in range(0, h - strip_h, strip_h):
        variances.append(np.var(gray[y:y+strip_h, :]))

    best = int(np.argmax(variances))
    center_y = best * strip_h + strip_h // 2

    # Вырезаем crop_ratio высоты изображения вокруг центра этикетки
    lh = int(h * crop_ratio)
    y1 = max(0, center_y - lh // 2)
    y2 = min(h, y1 + lh)
    if y2 >= h:
        y1 = max(0, h - lh)
        y2 = h

    return 0, y1, w, y2

# ── 3. Обработать один файл ──
def crop_label(filename):
    path = FOLDER / filename
    if not path.exists():
        print(f"  ⚠ Файл не найден: {filename}")
        return False

    out_path = FOLDER / ("label_" + filename)
    if out_path.exists():
        print(f"  ↩ Уже есть: label_{filename}")
        return True

    img = Image.open(path).convert("RGBA")
    arr = np.array(img)

    x1, y1, x2, y2 = find_label_box(arr)
    cropped = img.crop((x1, y1, x2, y2))

    # Приводим к единому размеру 400×200 (широкий формат этикетки)
    TARGET = (400, 200)
    cropped = cropped.resize(TARGET, Image.LANCZOS)
    cropped.save(out_path, "PNG")
    return True

# ── 4. Обновить HTML ──
def update_html(html):
    # Меняем шаблон карточки: для wine и bar используем label_ как превью
    old = "const thumb = i.img;"
    new = "const thumb = (sec === 'wine' || sec === 'bar') && i.img ? 'label_' + i.img : i.img;"

    if old in html:
        return html.replace(old, new)

    # Если переменной thumb нет — внедряем логику в шаблон карточки напрямую
    old2 = '? `<img src="${i.img}" alt="${i.name}" loading="lazy" onerror="this.src=\'${FALLBACK}\'" onclick="openLightbox(this.src)" style="cursor:zoom-in">`'
    new2 = '? `<img src="${(sec===\'wine\'||sec===\'bar\')&&i.img?\'label_\'+i.img:i.img}" alt="${i.name}" loading="lazy" onerror="this.onerror=null;this.src=\'${i.img||FALLBACK}\'" onclick="openLightbox(\'${i.img}\')" style="cursor:zoom-in">`'

    if old2 in html:
        return html.replace(old2, new2)

    print("  ⚠ Шаблон карточки не найден — HTML не изменён")
    return html

# ── Main ──
def main():
    html = HTML.read_text(encoding="utf-8")

    all_imgs = []
    for sec in ("bar", "wine"):
        imgs = extract_imgs(html, sec)
        print(f"\n{sec}: найдено {len(imgs)} изображений")
        all_imgs.extend(imgs)

    # Убираем дубли
    all_imgs = list(dict.fromkeys(all_imgs))
    print(f"\nВсего уникальных файлов: {len(all_imgs)}\n")

    done = errors = 0
    for i, fname in enumerate(all_imgs, 1):
        print(f"[{i}/{len(all_imgs)}] {fname}", end=" ... ", flush=True)
        if crop_label(fname):
            done += 1
            print("✓")
        else:
            errors += 1

    print(f"\n✅ Создано: {done} превью, ошибок: {errors}")

    print("\nОбновляю HTML...")
    new_html = update_html(html)
    if new_html != html:
        HTML.write_text(new_html, encoding="utf-8")
        print("✅ HTML обновлён")
    else:
        print("⚠ HTML не изменился — проверьте вручную")

if __name__ == "__main__":
    main()
