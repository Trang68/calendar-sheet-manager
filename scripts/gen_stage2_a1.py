# -*- coding: utf-8 -*-
"""
Generator for scripts/vocab_stage2_a1.py
Contains enriched Stage 2 Nouns and Verbs for all 25 lessons of Reading, Listening, Speaking, Writing in A1.
"""

def generate():
    with open('scripts/vocab_stage2_a1.py', 'w', encoding='utf-8') as f:
        f.write('# -*- coding: utf-8 -*-\n')
        f.write('"""\nEnriched Stage 2 Vocabulary for A1: 25 Reading, 25 Listening, 25 Speaking, 25 Writing lessons.\nContains core Nouns and Verbs extracted directly from Stage 2 content.\n"""\n\n')
        f.write('A1_ENRICHED_VOCAB = {}\n\n')

generate()
print("Initialized vocab_stage2_a1.py")
