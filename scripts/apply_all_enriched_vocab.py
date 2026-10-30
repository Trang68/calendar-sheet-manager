# -*- coding: utf-8 -*-
"""
Apply enriched vocabulary (>= 8 words, Nouns + Verbs from Stage 2) to all 300 lessons
across A1, A2, B1 for Reading, Listening, Speaking, and Writing.

STRICT SAFETY:
- Explicit assert ensures lvl != 'b2' and lvl != 'c1'.
- B2 and C1 are NEVER touched.
"""
import json
import os
import sys

LEVELS = ['a1', 'a2', 'b1']
SKILLS = ['reading', 'listening', 'speaking', 'writing']

applied_count = 0

for lvl in LEVELS:
    assert lvl not in ['b2', 'c1'], f"CRITICAL ERROR: Cannot touch level {lvl}!"
    master_path = f'scripts/vocab_master_{lvl}.json'
    with open(master_path, 'r', encoding='utf-8') as f:
        master_data = json.load(f)

    for skill in SKILLS:
        skill_data = master_data.get(skill, {})
        for lid in range(1, 26):
            lid_str = str(lid)
            assert lid_str in skill_data, f"Missing {lvl} {skill} {lid} in master"
            enriched_cats = skill_data[lid_str]

            lesson_path = f'public/data/{skill}/{lvl}/lesson-{lid:02d}.json'
            assert os.path.exists(lesson_path), f"File not found: {lesson_path}"

            with open(lesson_path, 'r', encoding='utf-8') as lf:
                lesson_json = json.load(lf)

            # Update vocabCategories
            lesson_json['vocabCategories'] = enriched_cats

            # Save back with pretty indent
            with open(lesson_path, 'w', encoding='utf-8') as lf:
                json.dump(lesson_json, lf, ensure_ascii=False, indent=2)

            applied_count += 1

print(f"SUCCESS: Successfully applied enriched vocabulary to {applied_count} lesson files!")
