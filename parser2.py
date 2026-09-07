import os
import time
from datetime import datetime
from common import (
    session_get, setup_logging, load_config, save_config,
    save_progress, load_progress, clear_progress, parse_search_url,
    process_vacancy_details, sort_vacancies_by_date, save_to_xlsx,
)

CONFIG_FILE = 'config.json'
LOG_FILE = 'parser2.log'
PROGRESS_FILE = 'progress2.json'

# ============================================================
# ЛОГИРОВАНИЕ (ОБНОВЛЕНО)
# ============================================================
# API ЗАПРОСЫ
# ============================================================
def fetch_page(token, params, page, per_page, logger, verbose):
    headers = {'Authorization': f'Bearer {token}', 'User-Agent': 'VacancyParser/1.0', 'Accept': 'application/json'}
    params['page'] = page
    params['per_page'] = min(per_page, 100)
    
    try:
        r = session_get('https://api.hh.ru/vacancies', logger, verbose, headers=headers, params=params)
        if verbose:
            logger.info(f"📡 GET {r.url}")
            logger.debug(f"📥 Ответ: статус {r.status_code}, длина {len(r.text)} символов")
        return r.json()
    except Exception as e:
        if verbose:
            logger.error(f"❌ Ошибка загрузки страницы {page}: {e}")
        return None

def fetch_details(token, vid, logger, verbose):
    headers = {'Authorization': f'Bearer {token}', 'User-Agent': 'VacancyParser/1.0', 'Accept': 'application/json'}
    url = f'https://api.hh.ru/vacancies/{vid}'
    try:
        r = session_get(url, logger, verbose, headers=headers)
        if verbose:
            logger.debug(f"📡 GET {url} → статус {r.status_code}")
        return r.json()
    except Exception as e:
        if verbose:
            logger.error(f"❌ Ошибка загрузки вакансии {vid}: {e}")
        return None

# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================
def main():
    print("=" * 60)
    print("  ПАРСЕР ВАКАНСИЙ HH.RU (ТОЛЬКО API)")
    print("=" * 60)
    
    cfg = load_config(sort_default=True)
    verbose = cfg.get('verbose', False)
    logger = setup_logging('hh_parser2', LOG_FILE, verbose)
    
    if verbose:
        logger.info("=== Запуск парсера (API only) ===")
        logger.info(f"Конфиг: {cfg}")
    
    token = cfg['api_token']
    if token == "ВАШ_ТОКЕН_СЮДА":
        print("❌ Заполните api_token в config.json")
        if verbose:
            logger.error("API токен не указан")
        return
    
    print(f"✅ Токен загружен")
    
    # ===== ПРОВЕРКА ПРОГРЕССА =====
    progress = load_progress(PROGRESS_FILE)
    resume = False
    
    if progress:
        print(f"\n📂 Найден файл прогресса: {PROGRESS_FILE}")
        print(f"   📊 Сохранено вакансий: {len(progress.get('vacancy_ids', []))}")
        print(f"   📄 Последняя страница: {progress.get('last_page', 0)}")
        
        choice = input("\n❓ Продолжить с места остановки? (y/n): ").strip().lower()
        if choice == 'y':
            resume = True
            print("✅ Продолжаем загрузку...")
        else:
            print("🗑️ Начинаем заново...")
            clear_progress(PROGRESS_FILE)
    
    # ===== НАСТРОЙКИ =====
    if resume and progress:
        url = progress.get('url', '')
        max_vacancies = progress.get('max_vacancies', 100)
        start_page = progress.get('last_page', 0) + 1
        end_page = progress.get('end_page', -1)
        sort_by_date = progress.get('sort_by_date', True)
        vacancy_ids = progress.get('vacancy_ids', [])
        vacancies_data = progress.get('vacancies_data', {})
        search_params = progress.get('search_params', {})
        
        print(f"\n📋 Настройки из прогресса:")
        print(f"   🔹 Макс. вакансий: {max_vacancies}")
        print(f"   🔹 Продолжаем со страницы: {start_page}")
        print(f"   🔹 Уже собрано ID: {len(vacancy_ids)}")
    else:
        end_page_display = cfg.get('end_page', -1)
        end_page_str = "до конца" if end_page_display == -1 else str(end_page_display)
        
        print("\n" + "-" * 60)
        print("📋 Настройки в config.json:")
        print(f"   🔹 Макс. вакансий: {cfg.get('max_vacancies', 100)}")
        print(f"   🔹 Стартовая страница: {cfg.get('start_page', 0)}")
        print(f"   🔹 Конечная страница: {end_page_str}")
        print(f"   🔹 Сортировка по дате: {'ДА' if cfg.get('sort_by_date') else 'НЕТ'}")
        print("-" * 60)
        
        use_config = input("\n❓ Использовать эти настройки? (y/n): ").strip().lower()
        
        if use_config == 'y':
            max_vacancies = cfg.get('max_vacancies', 100)
            start_page = cfg.get('start_page', 0)
            end_page = cfg.get('end_page', -1)
            sort_by_date = cfg.get('sort_by_date', True)
        else:
            print("\n📝 Ручной ввод параметров:")
            try:
                max_vacancies = int(input("   Максимум вакансий (Enter=100): ") or "100")
                start_page = int(input("   Начать со страницы (Enter=0): ") or "0")
                
                end_input = input("   До какой страницы (Enter=-1 = до конца): ").strip()
                end_page = int(end_input) if end_input.lstrip('-').isdigit() else -1
                
                sort_input = input("   Сортировать по дате? (y/n, Enter=y): ").strip().lower()
                sort_by_date = sort_input != 'n'
            except ValueError:
                print("❌ Введите корректные числа")
                return
            
            save = input("\n💾 Сохранить настройки в config.json? (y/n): ").strip().lower()
            if save == 'y':
                cfg['max_vacancies'] = max_vacancies
                cfg['start_page'] = start_page
                cfg['end_page'] = end_page
                cfg['sort_by_date'] = sort_by_date
                save_config(cfg)
        
        url = input("\n🔗 Ссылка поиска hh.ru: ").strip()
        if not url or 'hh.ru' not in url:
            print("❌ Некорректная ссылка")
            return
        
        search_params = parse_search_url(url)
        vacancy_ids = []
        vacancies_data = {}
    
    if verbose:
        logger.info(f"🔗 URL: {url}")
        logger.info(f"📋 Параметры: {search_params}")
    
    if not search_params and not resume:
        print("❌ Не удалось извлечь параметры из ссылки")
        return
    
    end_page_str = "до конца" if end_page == -1 else str(end_page)
    print(f"✅ Параметры: {list(search_params.keys())}")
    print(f"🚀 Сбор (API):")
    print(f"   📊 Лимит: {max_vacancies}")
    print(f"   📄 Страницы: {start_page} → {end_page_str}")
    print(f"   📅 Сортировка: {'ДА' if sort_by_date else 'НЕТ'}")
    print()
    
    # ===== ОСНОВНОЙ ЦИКЛ (API) =====
    page = start_page
    
    while True:
        if end_page != -1 and page > end_page:
            print(f"\n🛑 Достигнута конечная страница {end_page}.")
            if verbose:
                logger.info(f"Достигнута конечная страница {end_page}")
            break
        
        if len(vacancy_ids) >= max_vacancies:
            print(f"\n🛑 Достигнут лимит {max_vacancies}.")
            if verbose:
                logger.info(f"Достигнут лимит {max_vacancies}")
            break
        
        print(f"\n📄 Страница {page}...")
        
        data = fetch_page(token, search_params, page, cfg.get('per_page', 100), logger, verbose)
        
        if not data:
            print("   ❌ Не удалось получить данные")
            break
        
        items = data.get('items', [])
        total_found = data.get('found', 0)
        
        if page == start_page:
            print(f"   📊 Всего найдено в API: {total_found}")
        
        if not items:
            print("   ⚠️ Больше вакансий нет")
            break
        
        print(f"   📋 Найдено {len(items)} вакансий на странице")
        
        added = 0
        for item in items:
            vid = item['id']
            if vid not in vacancies_data and len(vacancy_ids) < max_vacancies:
                vacancy_ids.append(vid)
                added += 1
        
        print(f"   ➕ Добавлено: {added} | Всего ID: {len(vacancy_ids)}")
        
        print(f"   ⚡ Загрузка деталей...")
        loaded_count = 0
        for item in items:
            vid = item['id']
            if vid not in vacancies_data:
                details = fetch_details(token, vid, logger, verbose)
                if details:
                    vacancies_data[vid] = details
                    loaded_count += 1
                    if verbose:
                        logger.info(f"   ✅ Загружена: ID={vid}, Название={details.get('name', '-')[:60]}")
                time.sleep(0.1)
        
        print(f"   ✅ Загружено деталей: {loaded_count}")
        
        # 💾 СОХРАНЯЕМ ПРОГРЕСС
        progress_data = {
            'url': url,
            'search_params': search_params,
            'max_vacancies': max_vacancies,
            'start_page': start_page,
            'end_page': end_page,
            'sort_by_date': sort_by_date,
            'last_page': page,
            'vacancy_ids': vacancy_ids,
            'vacancies_data': vacancies_data,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        save_progress(PROGRESS_FILE, progress_data)
        print(f"   💾 Прогресс сохранён")
        
        time.sleep(0.5)
        page += 1
    
    if not vacancy_ids:
        print("❌ Вакансии не найдены")
        if verbose:
            logger.error("Вакансии не найдены")
        return
    
    vacancy_ids = vacancy_ids[:max_vacancies]
    
    remaining = [vid for vid in vacancy_ids if vid not in vacancies_data]
    if remaining:
        print(f"\n⚡ Загрузка оставшихся деталей ({len(remaining)})...")
        if verbose:
            logger.info(f"Загрузка оставшихся деталей: {len(remaining)}")
        for vid in remaining:
            details = fetch_details(token, vid, logger, verbose)
            if details:
                vacancies_data[vid] = details
                if verbose:
                    logger.info(f"   ✅ Загружена: ID={vid}, Название={details.get('name', '-')[:60]}")
            time.sleep(0.3)
    
    print(f"\n🔧 Обработка {len(vacancy_ids)} вакансий...")
    vacancies = []
    
    for i, vid in enumerate(vacancy_ids, 1):
        if vid in vacancies_data:
            vacancy_data = process_vacancy_details(vacancies_data[vid], i)
            if vacancy_data:
                vacancies.append(vacancy_data)
    
    if not vacancies:
        print("❌ Не удалось обработать вакансии")
        return
    
    print(f"\n✅ Собрано {len(vacancies)} вакансий")
    
    if sort_by_date:
        print("📅 Сортировка по дате...")
        vacancies = sort_vacancies_by_date(vacancies)
        for i, v in enumerate(vacancies, 1):
            v['num'] = i
        if verbose:
            logger.info("Вакансии отсортированы по дате")
    else:
        print("📋 Сохраняем порядок API")
        if verbose:
            logger.info("Порядок сохранён как в API")
    
    filename = f"hh_vacancies_api_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
    save_to_xlsx(vacancies, filename)
    
    print(f"📊 Размер: {os.path.getsize(filename)/1024:.1f} КБ")
    if verbose:
        print(f"📝 Логи: {LOG_FILE}")
    
    clear_progress()
    
    try:
        import subprocess
        subprocess.Popen(['explorer', os.path.dirname(os.path.abspath(filename))])
    except Exception:
        pass
    
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹ Прервано. Прогресс сохранён в progress2.json")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        print("💡 Прогресс сохранён в progress2.json")
        import traceback
        traceback.print_exc()
        input("\nНажмите Enter...")