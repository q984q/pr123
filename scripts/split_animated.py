# -*- coding: utf-8 -*-
"""Split the Animated deck into Drive-uploadable parts (valid pptx per part, source part boundaries)."""
import sys, os
sys.path.insert(0, "/tmp")
import build_all_decks as B

plan = B.plan_variant_cut("Animated", 72, B.PATCH_STAGES_CITY, [])
total = len(plan)
print(f"Animated plan: {total} slides")

# Source part boundaries (1-based, inclusive); last part absorbs custom slides 75-77
bounds = [(1,7),(8,18),(19,24),(25,34),(35,41),(42,48),(49,54),(55,61),(62,69),(70,73),(74,77)]
outdir = os.path.join(B.OUT, "SOKOL_Animated_RU_parts")
os.makedirs(outdir, exist_ok=True)
manifest = ["# Список частей — SOKOL_Animated_RU.pptx (витринная версия, 77 слайдов)", "",
            "Полная презентация разделена на части (лимит загрузки ~20 МБ, как в исходном комплекте).",
            "Нумерация слайдов сквозная (01–77) и совпадает с реестром REESTR_SOKOL_Animated.md.", ""]

for i, (a, b) in enumerate(bounds, start=1):
    sub = plan[a-1:b]
    out = os.path.join(outdir, f"SOKOL_Animated_RU__часть_{i:02d}__слайды_{a:03d}-{b:03d}.pptx")
    n, size, stripped, hits = B.build_deck(
        f"Animated_p{i}", sub, "Animated", False, "sokol", "SOKOL INDUSTRIAL PARK",
        out, idx_offset=a-1, total_override=total)
    print(f"  part {i:02d}: slides {a}-{b} -> {n} slides, {size/1e6:.1f} MB, patch hits {len(hits)}")
    manifest.append(f"| часть {i:02d} | слайды {a:03d}–{b:03d} | {os.path.basename(out)} | {size/1e6:.1f} МБ |")

with open(os.path.join(outdir, "Список_частей.md"), "w") as f:
    f.write("\n".join(manifest) + "\n")
print("manifest written")
