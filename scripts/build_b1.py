#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Builder script to generate full 25 lessons for Level B1 across Reading, Listening, Speaking, Writing.
"""
import json
import os

from gen_b1_reading import B1_READING_6_TO_25
from gen_b1_listening import B1_LISTENING_4_TO_25
from gen_b1_speaking import B1_SPEAKING_4_TO_25
from gen_b1_writing import B1_WRITING_4_TO_25

BASE_DIR = "public/data"

def load_existing_lessons(skill):
    dir_path = os.path.join(BASE_DIR, skill, "b1")
    lessons = []
    for i in range(1, 26):
        file_path = os.path.join(dir_path, f"lesson-{i:02d}.json")
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                lessons.append(json.load(f))
    return lessons

def build_b1_skill(skill, extra_lessons, badge, time_key):
    dir_path = os.path.join(BASE_DIR, skill, "b1")
    os.makedirs(dir_path, exist_ok=True)
    
    existing = load_existing_lessons(skill)
    existing_ids = {item["id"] for item in existing}
    
    # Merge existing and new
    all_lessons = list(existing)
    for extra in extra_lessons:
        if extra["id"] not in existing_ids:
            all_lessons.append(extra)
        else:
            # Overwrite or update
            for idx, item in enumerate(all_lessons):
                if item["id"] == extra["id"]:
                    all_lessons[idx] = extra
                    break
                    
    all_lessons.sort(key=lambda x: x["id"])
    
    # Write each lesson file
    manifest_items = []
    for item in all_lessons:
        lid = item["id"]
        fname = f"lesson-{lid:02d}.json"
        fpath = os.path.join(dir_path, fname)
        with open(fpath, "w", encoding="utf-8") as f:
            json.dump(item, f, ensure_ascii=False, indent=2)
            
        m_item = {
            "id": lid,
            "file": fname,
            "title": item["title"],
            "badge": badge,
            time_key: item.get(time_key, "20 phút"),
            "topic": item.get("topic", "")
        }
        if skill == "reading":
            m_item["questionCount"] = len(item.get("questions", []))
        elif skill == "listening":
            m_item["questionCount"] = len(item.get("questions", []))
        elif skill == "speaking":
            m_item["questionCount"] = len(item.get("shadowingPhrases", []))
        elif skill == "writing":
            m_item["targetWordRange"] = item.get("targetWordRange", [100, 130])
            
        manifest_items.append(m_item)
        
    manifest_path = os.path.join(dir_path, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest_items, f, ensure_ascii=False, indent=2)
        
    print(f"Built B1 {skill}: {len(all_lessons)} lessons, manifest updated.")

if __name__ == "__main__":
    build_b1_skill("reading", B1_READING_6_TO_25, "B1 Đọc hiểu", "readingTime")
    build_b1_skill("listening", B1_LISTENING_4_TO_25, "B1 Nghe hiểu", "listeningTime")
    build_b1_skill("speaking", B1_SPEAKING_4_TO_25, "B1 Nói", "speakingTime")
    build_b1_skill("writing", B1_WRITING_4_TO_25, "B1 Viết", "writingTime")
