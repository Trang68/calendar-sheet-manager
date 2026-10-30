# -*- coding: utf-8 -*-
"""
Applies curated vocabCategories to Listening, Speaking, and Writing lessons across A1, A2, and B1.
Never modifies B2 or Reading lessons (Reading lessons already have vocabCategories).
"""
import json
import os
import sys

# Ensure current directory is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import scripts.vocab_data_a1 as a1
import scripts.vocab_data_a2 as a2
import scripts.vocab_data_b1 as b1

data_map = {
    'a1': {
        'listening': a1.A1_LISTENING_VOCAB,
        'speaking': a1.A1_SPEAKING_VOCAB,
        'writing': a1.A1_WRITING_VOCAB
    },
    'a2': {
        'listening': a2.A2_LISTENING_VOCAB,
        'speaking': a2.A2_SPEAKING_VOCAB,
        'writing': a2.A2_WRITING_VOCAB
    },
    'b1': {
        'listening': b1.B1_LISTENING_VOCAB,
        'speaking': b1.B1_SPEAKING_VOCAB,
        'writing': b1.B1_WRITING_VOCAB
    }
}

def main():
    updated_count = 0
    for lvl, skills in data_map.items():
        assert lvl != 'b2', 'B2 must never be modified!'
        for skill, vocab_dict in skills.items():
            for lesson_id, vocab_cats in vocab_dict.items():
                fname = f'public/data/{skill}/{lvl}/lesson-{lesson_id:02d}.json'
                if not os.path.exists(fname):
                    raise FileNotFoundError(f'Missing file: {fname}')
                
                with open(fname, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                data['vocabCategories'] = vocab_cats
                
                with open(fname, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    f.write('\n')
                updated_count += 1

    print(f'Successfully applied vocabCategories to {updated_count} lessons.')

if __name__ == '__main__':
    main()
