#!/usr/bin/env python3
"""
Скрипт для:
1. Співставлення винятків з ignore_rules.csv з правилами в translation_rules.csv
2. Заповнення типів у translation_rules.csv
"""

import csv
import re

TRANSLATION_RULES_CSV = 'translation_rules.csv'
IGNORE_RULES_CSV = 'ignore_rules.csv'

def load_translation_rules():
    """Завантажити правила перекладу"""
    rules = []
    with open(TRANSLATION_RULES_CSV, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rules.append(row)
    return rules

def load_ignore_rules():
    """Завантажити винятки"""
    ignores = []
    with open(IGNORE_RULES_CSV, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            phrase = row.get('фраза', '').strip()
            if phrase:
                ignores.append(phrase)
    return ignores

def match_exceptions_to_rules(rules, ignores):
    """
    Співставити винятки з правилами.
    Логіка: виняток відноситься до правила, якщо містить один з українських коренів.
    """
    for rule in rules:
        ua_stems_str = rule.get('українські_корені', '').strip()
        if not ua_stems_str:
            continue
        
        ua_stems = [stem.strip() for stem in ua_stems_str.split(',')]
        matched_exceptions = []
        
        for exception in ignores:
            exception_lower = exception.lower()
            # Перевіряємо чи містить виняток один з українських коренів
            for ua_stem in ua_stems:
                if ua_stem and ua_stem.lower() in exception_lower:
                    matched_exceptions.append(exception)
                    break
        
        # Додаємо знайдені винятки до існуючих
        existing = rule.get('виключення', '').strip()
        if existing:
            existing_set = set(e.strip() for e in existing.split(','))
        else:
            existing_set = set()
        
        existing_set.update(matched_exceptions)
        
        if existing_set:
            rule['виключення'] = ','.join(sorted(existing_set))

def determine_type(rule):
    """
    Визначити тип правила на основі коментаря та контексту
    """
    comment = rule.get('коментар', '').lower()
    ru_stem = rule.get('російський_корінь', '').lower()
    
    # Ключові слова для визначення типу
    if 'суржик' in comment or 'калька' in comment:
        return 'Суржик'
    elif 'контекст' in comment or 'плутає' in comment or 'значення' in comment:
        return 'Помилка контексту'
    elif 'перекладач' in comment or 'програма' in comment or 'перекладає' in comment:
        return 'Помилка перекладу'
    
    # За ключовими словами в російському корені
    surzyk_stems = ['проводн', 'беспроводн', 'емкост', 'расход', 'ссылк', 'загрузк', 'полос']
    if ru_stem in surzyk_stems:
        return 'Суржик'
    
    # За замовчуванням - помилка перекладу
    return 'Помилка перекладу'

def fill_types(rules):
    """Заповнити типи для всіх правил"""
    for rule in rules:
        current_type = rule.get('тип', '').strip()
        if not current_type:
            rule['тип'] = determine_type(rule)

def save_translation_rules(rules):
    """Зберегти оновлені правила"""
    fieldnames = ['російський_корінь', 'українські_корені', 'виключення', 'тип', 'коментар']
    with open(TRANSLATION_RULES_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rules)

def main():
    print("Завантаження даних...")
    rules = load_translation_rules()
    ignores = load_ignore_rules()
    
    print(f"Завантажено {len(rules)} правил та {len(ignores)} винятків")
    
    print("\nСпівставлення винятків з правилами...")
    match_exceptions_to_rules(rules, ignores)
    
    print("Заповнення типів...")
    fill_types(rules)
    
    print("\nЗбереження оновлених правил...")
    save_translation_rules(rules)
    
    print("✓ Готово!")
    print("\nСтатистика:")
    
    # Підрахунок статистики
    with_exceptions = sum(1 for r in rules if r.get('виключення', '').strip())
    with_types = sum(1 for r in rules if r.get('тип', '').strip())
    
    types_count = {}
    for r in rules:
        t = r.get('тип', '').strip()
        if t:
            types_count[t] = types_count.get(t, 0) + 1
    
    print(f"  - Правил з винятками: {with_exceptions}/{len(rules)}")
    print(f"  - Правил з типами: {with_types}/{len(rules)}")
    print(f"  - Розподіл за типами:")
    for type_name, count in sorted(types_count.items()):
        print(f"    - {type_name}: {count}")

if __name__ == '__main__':
    main()
