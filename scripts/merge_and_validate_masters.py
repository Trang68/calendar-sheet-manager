# -*- coding: utf-8 -*-
"""
Merge Reading A2 and B1 into master files, then validate all three masters (A1, A2, B1).
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.reading_master_a2_b1 import reading_a2_all, reading_b1_all

# 1. Update A2 Master
with open('scripts/vocab_master_a2.json', 'r', encoding='utf-8') as f:
    a2_master = json.load(f)

# Convert integer keys of reading_a2_all to string keys
a2_master['reading'] = {str(k): v for k, v in reading_a2_all.items()}

with open('scripts/vocab_master_a2.json', 'w', encoding='utf-8') as f:
    json.dump(a2_master, f, ensure_ascii=False, indent=2)
print("Updated scripts/vocab_master_a2.json with 25 Reading lessons.")

# 2. Update B1 Master
with open('scripts/vocab_master_b1.json', 'r', encoding='utf-8') as f:
    b1_master = json.load(f)

b1_master['reading'] = {str(k): v for k, v in reading_b1_all.items()}

with open('scripts/vocab_master_b1.json', 'w', encoding='utf-8') as f:
    json.dump(b1_master, f, ensure_ascii=False, indent=2)
print("Updated scripts/vocab_master_b1.json with 25 Reading lessons.")

# 3. Comprehensive Validation
masters = {
    'a1': 'scripts/vocab_master_a1.json',
    'a2': 'scripts/vocab_master_a2.json',
    'b1': 'scripts/vocab_master_b1.json'
}

skills = ['reading', 'listening', 'speaking', 'writing']
all_ok = True
total_words = 0

for lvl, path in masters.items():
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f"\n=== Validating Level {lvl.upper()} ({path}) ===")
    lvl_words = 0
    for sk in skills:
        assert sk in data, f"Missing skill {sk} in {lvl}"
        lessons = data[sk]
        assert len(lessons) == 25, f"Expected 25 lessons for {lvl} {sk}, got {len(lessons)}"
        for lid_str, cats in lessons.items():
            lid = int(lid_str)
            assert len(cats) >= 2, f"Lesson {lvl} {sk} L{lid} has < 2 categories: {len(cats)}"
            lesson_words = sum(len(cat.get('rows', [])) for cat in cats)
            assert lesson_words >= 8, f"Lesson {lvl} {sk} L{lid} has < 8 words: {lesson_words}"
            lvl_words += lesson_words
            for cat in cats:
                assert 'title' in cat and len(cat['title']) > 0
                for r in cat.get('rows', []):
                    assert 'word' in r and len(r['word']) > 0
                    assert 'gram' in r and len(r['gram']) > 0
                    assert 'meaning' in r and len(r['meaning']) > 0
                    assert 'example' in r and len(r['example']) > 0
        sk_avg = lvl_words / (25 * (skills.index(sk) + 1))
        print(f"  Skill {sk:10s}: 25/25 lessons OK (>= 2 cats, >= 8 words each)")
    print(f"Level {lvl.upper()} Total Words: {lvl_words} across 100 lessons")
    total_words += lvl_words

print(f"\n==========================================")
print(f"ALL 3 MASTERS VALIDATED SUCCESSFULLY!")
print(f"Total Lessons: 300 | Total Curated Words: {total_words}")
print(f"Average Words/Lesson: {total_words / 300:.1f}")
print(f"==========================================")
