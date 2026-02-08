#!/usr/bin/env python3
"""Debug script to check slide layouts."""

import sys
from pptx import Presentation

if len(sys.argv) < 2:
    print("Usage: python3 debug_layout.py <file.pptx>")
    sys.exit(1)

pptx_path = sys.argv[1]
print(f"Analyzing: {pptx_path}\n")

prs = Presentation(pptx_path)

print("=== Available Layouts ===")
for i, layout in enumerate(prs.slide_layouts):
    print(f"  [{i}] {layout.name}")

print("\n=== Slide Masters and Their Layouts ===")
for master_idx, master in enumerate(prs.slide_masters):
    print(f"Master {master_idx}: {master.name}")
    for i, layout in enumerate(master.slide_layouts):
        print(f"  [{i}] {layout.name}")

print("\n=== Slides and Their Layouts ===")
for slide_idx, slide in enumerate(prs.slides, 1):
    layout = slide.slide_layout
    layout_name = layout.name
    master = layout.slide_master
    master_name = master.name if master else "Unknown"

    # Find the layout index in presentation's slide_layouts
    layout_index = None
    for i, l in enumerate(prs.slide_layouts):
        if l.name == layout_name:
            layout_index = i
            break

    # Find the layout index within its master
    master_layout_index = None
    if master:
        for i, l in enumerate(master.slide_layouts):
            if l.name == layout_name:
                master_layout_index = i
                break

    print(f"Slide {slide_idx}: {layout_name} (prs index: {layout_index}, master: {master_name}, master index: {master_layout_index})")
