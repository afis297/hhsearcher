import os
import re
import time
import http.cookiejar as cookielib          # для работы с cookies
from datetime import datetime
from urllib.parse import urlparse, parse_qs, urlencode
from common import (
    session_get, setup_logging, load_config, save_config,
    save_progress, load_progress, clear_progress, parse_search_url,
    process_vacancy_details, sort_vacancies_by_date, save_to_xlsx,
)

CONFIG_FILE = 'config.json'
LOG_FILE = 'parser.log'
PROGRESS_FILE = 'progress.json'

# ============================================================
# COOKIES
# ============================================================
def load_cookies_from_file(cookie_file='cookies.txt'):
    cj = cookielib.MozillaCookieJar()
    try:
        cj.load(cookie_file, ignore_expires=True, ignore_discard=True)
        return cj
    except Exception as e:
        print(f"⚠️ Не удалось загрузить cookies: {e}")
        return None

# ============================================================
# ФИЛЬТРАЦИЯ ВАКАНСИЙ СО СТАТУСОМ "СОБЕСЕДОВАНИЕ"
# ============================================================
def is_near_interview(html, vid, context_chars=300):
    """
    Проверяет, находится ли слово 'Собеседование' в окрестности ссылки на вакансию.
    Возвращает True, если рядом есть статус "Собеседование" – такую вакансию нужно исключить.
    """
    pattern = f'/vacancy/{vid}'
    pos = html.find(pattern)
    if pos == -1:
        return False  # не найдено – пропускаем (не исключаем)
    
    start = max(0, pos - context_chars)
    end = min(len(html), pos + len(pattern) + context_chars)
    context = html[start:end]
    
    # Ищем слово "Собеседование" (с учётом разных падежей, но для простоты – только в именительном)
    if re.search(r'Собеседование', context, re.IGNORECASE):
        return True
    return False

# ============================================================
# ПАРСИНГ HTML СТРАНИЦЫ (С ФИЛЬТРАЦИЕЙ)
# ============================================================
def parse_html_page(url, page, logger, verbose, cookies=None, exclude_interview=True):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
    }
    
    parsed = urlparse(url)
    params = parse_qs(parsed.query)
    params['page'] = [str(page)]
    
    new_query = urlencode(params, doseq=True)
    full_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}?{new_query}"
    
    if verbose:
        logger.info(f"🌐 Загрузка HTML: {full_url}")
    
    try:
        response = session_get(full_url, logger, verbose, headers=headers, cookies=cookies)
        html = response.text
        
        # Все ID вакансий
        all_ids = re.findall(r'/vacancy/(\d+)', html)
        unique_ids = list(dict.fromkeys(all_ids))
        
        if exclude_interview:
            # Фильтруем те, рядом с которыми есть "Собеседование"
            filtered_ids = []
            for vid in unique_ids:
                if not is_near_interview(html, vid):
                    filtered_ids.append(vid)
            unique_ids = filtered_ids
        
        if verbose:
            logger.info(f"📋 Страница {page}: найдено {len(unique_ids)} вакансий (после фильтрации)")
            for idx, vid in enumerate(unique_ids, 1):
                logger.info(f"   [{idx:3d}] ID вакансии: {vid}")
        
        return unique_ids
        
    except Exception as e:
        if verbose:
            logger.error(f"❌ Ошибка загрузки HTML страницы {page}: {e}")
        return []

# ============================================================
# API ЗАПРОСЫ
# ============================================================
def fetch_vacancy_details(token, vid, logger, verbose):
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

# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================
def main():
    print("=" * 60)
    print("  ПАРСЕР ВАКАНСИЙ HH.RU (HTML + API)")
    print("=" * 60)
    
    cfg = load_config()
    verbose = cfg.get('verbose', False)
    logger = setup_logging('hh_parser', LOG_FILE, verbose)
    
    if verbose:
        logger.info("=== Запуск парсера ===")
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
        print(f"   ⏰ Время сохранения: {progress.get('timestamp', 'неизвестно')}")
        
        choice = input("\n❓ Продолжить с места остановки? (y/n): ").strip().lower()
        if choice == 'y':
            resume = True
            print("✅ Продолжаем загрузку...")
            if verbose:
                logger.info("Режим: продолжение с места остановки")
        else:
            print("🗑️ Начинаем заново...")
            clear_progress(PROGRESS_FILE)
            if verbose:
                logger.info("Режим: новая загрузка")
    
    # ===== НАСТРОЙКИ =====
    if resume and progress:
        url = progress.get('url', '')
        max_vacancies = progress.get('max_vacancies', 100)
        start_page = progress.get('last_page', 0) + 1
        end_page = progress.get('end_page', -1)
        sort_by_date = progress.get('sort_by_date', False)
        
        all_vacancy_ids = progress.get('vacancy_ids', [])
        vacancies_data = progress.get('vacancies_data', {})
        
        print(f"\n📋 Настройки из прогресса:")
        print(f"   🔹 Макс. вакансий: {max_vacancies}")
        print(f"   🔹 Продолжаем со страницы: {start_page}")
        print(f"   🔹 Уже собрано ID: {len(all_vacancy_ids)}")
        print(f"   🔹 Загружено деталей: {len(vacancies_data)}")
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
            sort_by_date = cfg.get('sort_by_date', False)
            print("✅ Используем настройки из config.json")
            if verbose:
                logger.info("Режим: настройки из конфига")
        else:
            print("\n📝 Ручной ввод параметров:")
            try:
                max_vacancies = int(input("   Максимум вакансий (Enter=100): ") or "100")
                start_page = int(input("   Начать со страницы (Enter=0): ") or "0")
                
                end_input = input("   До какой страницы (Enter=-1 = до конца): ").strip()
                end_page = int(end_input) if end_input.lstrip('-').isdigit() else -1
                
                sort_input = input("   Сортировать по дате? (y/n, Enter=n): ").strip().lower()
                sort_by_date = sort_input == 'y'
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
                print("✅ Настройки сохранены")
                if verbose:
                    logger.info(f"Настройки сохранены в {CONFIG_FILE}")
            
            if verbose:
                logger.info(f"Режим: ручной ввод (max={max_vacancies}, start={start_page}, end={end_page})")
        
        url = input("\n🔗 Ссылка поиска hh.ru: ").strip()
        if not url or 'hh.ru' not in url:
            print("❌ Некорректная ссылка")
            return
        
        all_vacancy_ids = []
        vacancies_data = {}
    
    if verbose:
        logger.info(f"🔗 Исходный URL: {url}")
    
    search_params = parse_search_url(url)
    if verbose:
        logger.info(f"📋 Параметры для API: {search_params}")
    
    if not search_params:
        print("ℹ️ Параметры фильтрации не обнаружены (возможно, это страница переговоров). Продолжаем без фильтров.")
    
    # ===== ЗАГРУЗКА COOKIES =====
    cookies = None
    if os.path.exists('cookies.txt'):
        cookies = load_cookies_from_file('cookies.txt')
        if cookies:
            print("🍪 Cookies загружены из cookies.txt")
        else:
            print("⚠️ Файл cookies.txt найден, но не удалось загрузить")
    else:
        print("ℹ️ Файл cookies.txt не найден. Для страниц переговоров требуется авторизация.")
    
    end_page_str = "до конца" if end_page == -1 else str(end_page)
    print(f"✅ Параметры извлечены: {list(search_params.keys()) if search_params else 'нет параметров'}")
    print(f"🚀 Начинаем сбор:")
    print(f"   📊 Лимит вакансий: {max_vacancies}")
    print(f"   📄 Страницы: {start_page} → {end_page_str}")
    print(f"   📅 Сортировка: {'ДА' if sort_by_date else 'НЕТ (порядок как на сайте)'}")
    print()
    
    # ===== ОСНОВНОЙ ЦИКЛ =====
    page = start_page
    
    while True:
        if end_page != -1 and page > end_page:
            print(f"\n🛑 Достигнута конечная страница {end_page}. Остановка.")
            if verbose:
                logger.info(f"Достигнута конечная страница {end_page}")
            break
        
        if len(all_vacancy_ids) >= max_vacancies:
            print(f"\n🛑 Достигнут лимит вакансий {max_vacancies}. Остановка.")
            if verbose:
                logger.info(f"Достигнут лимит вакансий {max_vacancies}")
            break
        
        print(f"\n📄 Страница {page}...")
        
        # Передаём cookies и включаем фильтрацию собеседований
        page_ids = parse_html_page(url, page, logger, verbose, cookies=cookies, exclude_interview=True)
        
        if not page_ids:
            print("   ⚠️ Больше вакансий нет или ошибка загрузки")
            break
        
        print(f"   📋 Найдено {len(page_ids)} вакансий (после исключения собеседований)")
        
        added = 0
        for vid in page_ids:
            if vid not in vacancies_data and len(all_vacancy_ids) < max_vacancies:
                all_vacancy_ids.append(vid)
                added += 1
        
        if added == 0:
            print("   ℹ️ На странице нет новых вакансий. Достигнут конец списка.")
            break
        
        print(f"   ➕ Добавлено новых: {added}")
        print(f"   📊 Всего ID: {len(all_vacancy_ids)}")
        
        print(f"   ⚡ Загрузка деталей через API...")
        
        loaded_count = 0
        for vid in page_ids:
            if vid not in vacancies_data:
                details = fetch_vacancy_details(token, vid, logger, verbose)
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
            'max_vacancies': max_vacancies,
            'start_page': start_page,
            'end_page': end_page,
            'sort_by_date': sort_by_date,
            'last_page': page,
            'vacancy_ids': all_vacancy_ids,
            'vacancies_data': vacancies_data,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        save_progress(PROGRESS_FILE, progress_data)
        print(f"   💾 Прогресс сохранён")
        
        time.sleep(1)
        page += 1
    
    if not all_vacancy_ids:
        print("❌ Вакансии не найдены")
        if verbose:
            logger.error("Вакансии не найдены")
        return
    
    all_vacancy_ids = all_vacancy_ids[:max_vacancies]
    
    remaining = [vid for vid in all_vacancy_ids if vid not in vacancies_data]
    if remaining:
        print(f"\n⚡ Загрузка оставшихся деталей ({len(remaining)})...")
        if verbose:
            logger.info(f"Загрузка оставшихся деталей: {len(remaining)}")
        for vid in remaining:
            details = fetch_vacancy_details(token, vid, logger, verbose)
            if details:
                vacancies_data[vid] = details
                if verbose:
                    logger.info(f"   ✅ Загружена: ID={vid}, Название={details.get('name', '-')[:60]}")
            time.sleep(0.3)
    
    print(f"\n🔧 Обработка {len(all_vacancy_ids)} вакансий...")
    vacancies = []
    
    for i, vid in enumerate(all_vacancy_ids, 1):
        if vid in vacancies_data:
            vacancy_data = process_vacancy_details(vacancies_data[vid], i)
            if vacancy_data:
                vacancies.append(vacancy_data)
    
    if not vacancies:
        print("❌ Не удалось обработать вакансии")
        return
    
    print(f"\n✅ Собрано {len(vacancies)} вакансий")
    
    if sort_by_date:
        print("📅 Сортировка по дате публикации...")
        vacancies = sort_vacancies_by_date(vacancies)
        for i, v in enumerate(vacancies, 1):
            v['num'] = i
        if verbose:
            logger.info("Вакансии отсортированы по дате")
    else:
        print("📋 Сохраняем порядок как на сайте (без сортировки)")
        if verbose:
            logger.info("Порядок сохранён как на сайте")
    
    filename = f"hh_vacancies_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
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
        print("\n\n⏹ Прервано. Прогресс сохранён в progress.json")
        print("💡 При следующем запуске можно продолжить с места остановки")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        print("💡 Прогресс сохранён в progress.json")
        import traceback
        traceback.print_exc()
        input("\nНажмите Enter...")