import os
import sys
import subprocess

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def print_menu():
    clear_screen()
    print("=" * 60)
    print("          ПАРСЕР ВАКАНСИЙ HH.RU + OLLAMA")
    print("=" * 60)
    print("\nВыберите действие:")
    print("  1. 🌐 Парсить через HTML + API")
    print("  2. 🚀 Парсить только через API")
    print("  3. 🔄 Объединить файлы XLSX")
    print("  4. 🤖 Обработать вакансии через Ollama")
    print("  5. 🚪 Выход")
    print("\n" + "=" * 60)

def check_file(filename):
    full_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    if not os.path.exists(full_path):
        print(f"\n❌ Файл {filename} не найден!")
        print(f"   Искал здесь: {full_path}")
        print(f"   Поместите {filename} в ту же папку, что и main.py")
        input("\nНажмите Enter для возврата в меню...")
        return False
    return True

def run_script(script_name):
    if not check_file(script_name):
        return
    print(f"\n⏳ Запуск {script_name}...\n")
    print("-" * 60)
    try:
        result = subprocess.run(
            [sys.executable, script_name],
            cwd=os.path.dirname(os.path.abspath(__file__)) or '.'
        )
        print("-" * 60)
        if result.returncode == 0:
            print(f"\n✅ {script_name} завершён успешно")
        else:
            print(f"\n⚠️ {script_name} завершился с кодом {result.returncode}")
    except Exception as e:
        print(f"\n❌ Ошибка запуска: {e}")

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__)) or '.'
    print(f"📂 Рабочая папка: {script_dir}\n")
    # Проверка наличия всех скриптов при старте (опционально)
    required = ['parser.py', 'parser2.py', 'merger.py', 'ollama_processing.py']
    missing = [f for f in required if not os.path.exists(os.path.join(script_dir, f))]
    if missing:
        print("⚠️ Внимание! Отсутствуют следующие скрипты:")
        for m in missing:
            print(f"   - {m}")
        print("Некоторые функции могут быть недоступны.\n")
    while True:
        print_menu()
        choice = input("\nВведите номер (1-5): ").strip()
        if choice == '1':
            run_script('parser.py')
            input("\nНажмите Enter для возврата в меню...")
        elif choice == '2':
            run_script('parser2.py')
            input("\nНажмите Enter для возврата в меню...")
        elif choice == '3':
            run_script('merger.py')
            input("\nНажмите Enter для возврата в меню...")
        elif choice == '4':
            run_script('ollama_processing.py')
            input("\nНажмите Enter для возврата в меню...")
        elif choice == '5':
            print("\n🔴 До свидания!")
            break
        else:
            print("\n🔴 Неверный выбор. Попробуйте снова.")
            input("Нажмите Enter...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹ Прервано пользователем")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()
        input("\nНажмите Enter для выхода...")