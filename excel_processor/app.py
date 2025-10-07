from flask import Flask, render_template, request, send_file, jsonify, url_for
import pandas as pd
import io
from datetime import datetime, timedelta
import logging
import gc
import os
from werkzeug.utils import secure_filename

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB limit
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Создаем папку для загрузок
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Увеличиваем лимиты pandas для больших файлов
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Сроки хранения по статусам
expiration_days = {
    "в заказе": {
        "PTE": 0, "PTS": 0, "DJR": 0, "ASP": 0, "WIJ": 0, "PBI": 0, "CTS": 3, "SPM": 4,
        "WDC": 0, "WJR": 0, "AIP": 0, "SDW": 0, "SGP": 0, "GIS": 0, "RWP": 0, "SDU": 0,
        "UBG": 0, "ADS": 21, "RSP": 0, "RNR": 21, "SDO": 0, "WSR": 0, "RME": 0, "WSC": 0,
        "SDL": 0, "WPF": 0, "WPT": 0, "WPU": 0, "RND": 0, "WDV": 0, "TWR": 0, "RNP": 21,
        "WWE": 0, "SFP": 0, "NAP": 4, "UDG": 0, "SAS": 8, "MTT": 0, "AGM": 0, "TMM": 8,
        "SMS": 7, "SMC": 7, "SAP": 4, "SSI": 4, "APC": 4, "SPS": 4, "SHC": 4, "PSC": 4,
        "ASI": 4, "BAP": 0, "APG": 4, "WSF": 4, "WMI": 4, "FSF": 4, "PIS": 3, "USD": 3,
        "PSR": 3, "LGR": 1
    },
    "без заказа": {
        "PTE": 60, "PTS": 38, "DJR": 38, "ASP": 25, "WIJ": 25, "PBI": 25, "CTS": 25,
        "SPM": 25, "WDC": 21, "WJR": 21, "AIP": 21, "SDW": 21, "SGP": 21, "GIS": 21,
        "RWP": 21, "SDU": 21, "UBG": 21, "ADS": 21, "RSP": 21, "RNR": 21, "SDO": 21,
        "WSR": 21, "RME": 21, "WSC": 21, "SDL": 21, "WPF": 21, "WPT": 21, "WPU": 21,
        "RND": 21, "WDV": 21, "TWR": 21, "RNP": 21, "WWE": 14, "SFP": 14, "NAP": 14,
        "UDG": 10, "SAS": 8, "MTT": 8, "AGM": 8, "TMM": 8, "SMS": 7, "SMC": 7, "SAP": 4,
        "SSI": 4, "APC": 4, "SPS": 4, "SHC": 4, "PSC": 4, "ASI": 4, "BAP": 4, "APG": 4,
        "WSF": 4, "WMI": 4, "FSF": 4, "PIS": 3, "USD": 3, "PSR": 3, "LGR": 1
    }
}

def allowed_file(filename):
    """Проверяем что файл Excel"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html')

def calculate_days_until_expiration(create_date, status, order_type):
    """Рассчитывает сколько дней осталось до списания"""
    try:
        # Парсим дату создания
        if isinstance(create_date, str):
            date_part = create_date.split()[0]
            formats_to_try = ['%Y-%m-%d', '%d.%m.%Y', '%Y.%m.%d', '%d/%m/%Y']
            parsed = False
            
            for fmt in formats_to_try:
                try:
                    create_date = datetime.strptime(date_part, fmt)
                    parsed = True
                    break
                except ValueError:
                    continue
            
            if not parsed:
                return 999
                
        elif isinstance(create_date, datetime):
            pass
        else:
            return 999
            
        # Получаем срок хранения для статуса
        status_clean = str(status).strip()
        expiration_days_count = expiration_days[order_type].get(status_clean, 0)
        
        # Рассчитываем дату списания
        expiration_date = create_date + timedelta(days=expiration_days_count)
        
        # Рассчитываем сколько дней осталось
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        expiration_date = expiration_date.replace(hour=0, minute=0, second=0, microsecond=0)
        create_date = create_date.replace(hour=0, minute=0, second=0, microsecond=0)
        
        days_left = (expiration_date - today).days
        
        return days_left
        
    except Exception as e:
        return 999

def modify_names_for_sc(df):
    """Заменяет наименование на значение из MX для СЦ"""
    if 'MX' not in df.columns or 'Наименование' not in df.columns:
        return df
    
    modified_count = 0
    
    for idx, row in df.iterrows():
        mx_value = str(row['MX'])
        current_name = str(row['Наименование'])
        
        # Заменяем наименование на значение из MX для трех конкретных случаев
        if 'Выход с сортировки СЦ' in mx_value:
            df.at[idx, 'Наименование'] = mx_value
            modified_count += 1
        elif 'Котовск_Буфер' in mx_value:
            df.at[idx, 'Наименование'] = mx_value
            modified_count += 1
        elif 'Принято на ворота' in mx_value:
            df.at[idx, 'Наименование'] = mx_value
            modified_count += 1
    
    print(f"✅ Заменено наименований на MX для СЦ: {modified_count}")
    return df

def apply_mx_filters(df, processing_type):
    """Применяем фильтры MX с оптимизацией"""
    if 'MX' not in df.columns:
        return df
    
    # Общий паттерн для удаления в ЛЮБОМ режиме
    patterns_to_remove_always = [
        'Выход с сортировки СЦ Котовск КС31'
    ]
    
    # Всегда удаляем Выход с сортировки СЦ Котовск КС31
    mask_always = df['MX'].astype(str).str.contains('|'.join(patterns_to_remove_always), na=False)
    removed_count = mask_always.sum()
    df = df[~mask_always]
    print(f"📊 После удаления общих паттернов: удалено {removed_count}, осталось {len(df)} строк")
    
    if processing_type == 'potok':
        # ПОТОК - удаляем дополнительные значения из столбца MX
        patterns_to_remove_potok = [
            'Выход с сортировки СЦ',
            'Выход с сортировки Предсорт СЦ', 
            'Предсорт СЦ',
            'СЦ Котовск КГТ',
            'Котовск_Буфер',
            'Буфер Предсорта СЦ',
            'Принято на ворота',
            'Упаковка ПСБ'
        ]
        # Удаляем строки, где MX содержит указанные паттерны
        mask = df['MX'].astype(str).str.contains('|'.join(patterns_to_remove_potok), na=False)
        removed_count = mask.sum()
        df = df[~mask]
        print(f"📊 После фильтрации MX для Потока: удалено {removed_count}, осталось {len(df)} строк")
        
    elif processing_type == 'sc':
        # СЦ - оставляем только определенные значения МХ
        sc_patterns = [
            'Выход с сортировки СЦ',
            'Выход с сортировки Предсорт СЦ', 
            'Предсорт СЦ',
            'СЦ Котовск КГТ',
            'СЦ Котовск КС',
            'Котовск_Буфер',
            'Буфер Предсорта СЦ',
            'Принято на ворота'
        ]
        
        # Сначала оставляем только разрешенные паттерны
        mask_keep = df['MX'].astype(str).str.contains('|'.join(sc_patterns), na=False)
        df_temp = df[mask_keep]
        
        # Теперь фильтруем Котовск_Буфер: оставляем только с цифрами после "в"
        def filter_kotovsk_buffer(mx_value):
            mx_str = str(mx_value)
            if 'Котовск_Буфер в' in mx_str:
                parts = mx_str.split('Котовск_Буфер в')
                if len(parts) > 1:
                    after_in = parts[1].strip()
                    if any(char.isdigit() for char in after_in):
                        return True
                    else:
                        return False
            elif 'Котовск_Буфер' in mx_str and ' в ' not in mx_str:
                return True
            return True
        
        mask_kotovsk = df_temp['MX'].apply(filter_kotovsk_buffer)
        df = df_temp[mask_kotovsk]
        
        removed_count = len(mask_keep) - len(df)
        print(f"📊 После фильтрации MX для СЦ: удалено {removed_count}, осталось {len(df)} строк")
        
        # Заменяем наименования на MX для СЦ
        print("🎯 Начинаю замену наименований на MX для СЦ...")
        df = modify_names_for_sc(df)
    
    return df

def add_expiration_column(df, order_type):
    """Добавляем столбец 'Осталось до списания'"""
    if 'Дата создания' not in df.columns or 'Товар' not in df.columns:
        return df
    
    # Читаем оригинальный файл для маппинга статусов
    original_df = df.copy()
    status_map = original_df.set_index('Товар')['Статус товара'].to_dict()
    
    def calculate_expiration_info(row):
        try:
            create_date = row['Дата создания']
            tovar = row['Товар']
            
            # Берем статус ПО Товар (уникальный идентификатор)
            status = status_map.get(tovar, '')
            
            # Парсим дату создания
            if isinstance(create_date, str):
                datetime_formats = [
                    '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M',
                    '%Y.%m.%d %H:%M:%S', '%Y.%m.%d %H:%M', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M'
                ]
                
                parsed_with_time = False
                for fmt in datetime_formats:
                    try:
                        create_date = datetime.strptime(create_date, fmt)
                        parsed_with_time = True
                        break
                    except ValueError:
                        continue
                
                if not parsed_with_time:
                    date_part = create_date.split()[0]
                    date_formats = ['%Y-%m-%d', '%d.%m.%Y', '%Y.%m.%d', '%d/%m/%Y']
                    for fmt in date_formats:
                        try:
                            create_date = datetime.strptime(date_part, fmt)
                            break
                        except ValueError:
                            continue
                    else:
                        return "Ошибка даты"
                        
            elif isinstance(create_date, datetime):
                pass
            else:
                return "Неизвестный формат даты"
            
            # Получаем срок хранения
            expiration_days_count = expiration_days[order_type].get(str(status).strip(), 0)
            
            # Рассчитываем дату списания
            expiration_date = create_date + timedelta(days=expiration_days_count)
            
            # Рассчитываем сколько осталось
            time_left = expiration_date - datetime.now()
            days = time_left.days
            hours = time_left.seconds // 3600
            
            # Логика отображения
            if days < 0:
                result = f"ПРОСРОЧЕНО ({abs(days)} дн)"
            elif days == 0:
                if hours <= 0:
                    result = "СПИСАНИЕ"
                else:
                    result = f"Сегодня ({hours} ч)"
            else:
                result = f"{days} дн {hours} ч"
            
            return result
                
        except Exception as e:
            return "Ошибка расчета"
    
    # Добавляем столбец
    df['Осталось до списания'] = df.apply(calculate_expiration_info, axis=1)
    print(f"✅ Добавлен столбец 'Осталось до списания'")
    
    return df

def sort_by_priority(df):
    """Сортировка по приоритету списания"""
    if 'Осталось до списания' not in df.columns:
        return df
    
    print("🎯 Начинаю сортировку по приоритету списания...")
    
    def get_sort_priority(expiration_text):
        """Определяет приоритет для сортировки"""
        if isinstance(expiration_text, str):
            if expiration_text.startswith('ПРОСРОЧЕНО'):
                try:
                    days = int(expiration_text.split('(')[1].split('дн')[0].strip())
                    return 100 - days
                except:
                    return 50
            elif expiration_text == "СПИСАНИЕ":
                return 200
            elif expiration_text.startswith('Сегодня'):
                try:
                    if '(' in expiration_text:
                        hours = int(expiration_text.split('(')[1].split('ч')[0].strip())
                        return 300 + hours
                    else:
                        return 300
                except:
                    return 300
            elif 'дн' in expiration_text:
                try:
                    parts = expiration_text.split('дн')
                    days = int(parts[0].strip())
                    if len(parts) > 1 and 'ч' in parts[1]:
                        hours = int(parts[1].split('ч')[0].strip())
                    else:
                        hours = 0
                    return 1000 + days * 100 + hours
                except:
                    return 9999
        return 9999
    
    # Добавляем столбец с приоритетом
    df['_sort_priority'] = df['Осталось до списания'].apply(get_sort_priority)
    
    # Сортируем по приоритету
    df = df.sort_values('_sort_priority')
    
    # Удаляем временный столбец
    df = df.drop(columns=['_sort_priority'])
    
    print(f"✅ Отсортировано строк: {len(df)}")
    return df

def generate_summary(df, order_type, processing_type):
    """Генерация текстовой сводки"""
    # Основная статистика
    total_products = len(df)
    
    # Кол-во одинаковых гофр (дублей)
    if 'Гофра' in df.columns:
        gofra_counts = df['Гофра'].value_counts()
        duplicate_gofras = len(gofra_counts[gofra_counts > 1])
    else:
        duplicate_gofras = 0
    
    # Общая стоимость
    total_cost = df['Стоимость'].sum() if 'Стоимость' in df.columns else 0
    
    # Дополнительная статистика для СЦ
    if processing_type == 'sc' and 'Осталось до списания' in df.columns:
        # Просроченные товары
        expired_mask = df['Осталось до списания'].astype(str).str.startswith('ПРОСРОЧЕНО')
        expired_count = expired_mask.sum()
        expired_cost = df[expired_mask]['Стоимость'].sum() if expired_count > 0 else 0
        
        # Товары для списания сегодня
        today_mask = df['Осталось до списания'].astype(str).str.startswith('Сегодня')
        today_count = today_mask.sum()
        today_cost = df[today_mask]['Стоимость'].sum() if today_count > 0 else 0
        
        # Товары для списания за 1 день
        one_day_mask = df['Осталось до списания'].astype(str).str.contains('1 дн')
        one_day_count = one_day_mask.sum()
        one_day_cost = df[one_day_mask]['Стоимость'].sum() if one_day_count > 0 else 0
        
        # Товары для списания за 2 дня
        two_days_mask = df['Осталось до списания'].astype(str).str.contains('2 дн')
        two_days_count = two_days_mask.sum()
        two_days_cost = df[two_days_mask]['Стоимость'].sum() if two_days_count > 0 else 0
        
        summary = f"""📊 <b>СВОДНАЯ ПО ОБРАБОТАННОЙ ТАБЛИЦЕ:</b><br><br>
📦 <b>Тип заказа:</b> {order_type}<br>
🏭 <b>Тип обработки:</b> СЦ<br><br>
📈 <b>Статистика:</b><br>
• Кол-во товаров: {total_products} шт<br>
• Кол-во одинаковых гофр: {duplicate_gofras} шт<br>
• Общая стоимость: {total_cost:,.0f} руб<br><br>
⏰ <b>Сроки списания:</b><br>
• Просрочено: {expired_count} шт - {expired_cost:,.0f} руб<br>
• Списание сегодня: {today_count} шт - {today_cost:,.0f} руб<br>
• Списание через 1 день: {one_day_count} шт - {one_day_cost:,.0f} руб<br>
• Списание через 2 дня: {two_days_count} шт - {two_days_cost:,.0f} руб"""
            
    else:
        # Стандартная сводка для Потока
        summary = f"""📊 <b>СВОДКА ПО ОБРАБОТАННОЙ ТАБЛИЦЕ:</b><br><br>
📦 <b>Тип заказа:</b> {order_type}<br>
🏭 <b>Тип обработки:</b> {'СЦ' if processing_type == 'sc' else 'Поток'}<br><br>
📈 <b>Статистика:</b><br>
• Кол-во товаров: {total_products} шт<br>
• Кол-во одинаковых гофр: {duplicate_gofras} шт<br>
• Общая стоимость: {total_cost:,.0f} руб<br>"""

    return summary

@app.route('/process', methods=['POST'])
def process_file():
    """Обработка загруженного файла"""
    try:
        # Проверяем что файл был отправлен
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Файл не выбран'})
        
        file = request.files['file']
        
        # Проверяем что файл имеет имя
        if file.filename == '':
            return jsonify({'success': False, 'error': 'Файл не выбран'})
        
        # Проверяем формат файла
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Разрешены только файлы Excel (.xlsx, .xls)'})
        
        # Получаем параметры из формы
        order_type = request.form.get('order_type', 'в заказе')
        processing_type = request.form.get('processing_type', 'sc')
        
        print(f"📁 Обрабатываю файл: {file.filename}")
        print(f"📦 Тип заказа: {order_type}")
        print(f"🏭 Тип обработки: {processing_type}")
        
        # Читаем файл
        df = pd.read_excel(
            file,
            engine='openpyxl',
            dtype={
                'Статус товара': 'category',
                'Гофра': 'string',
                'Товар': 'string', 
                'Наименование': 'string',
                'MX': 'string'
            }
        )
        
        print(f"✅ Файл прочитан: {len(df)} строк, {len(df.columns)} колонок")
        
        # Применяем фильтрацию по срокам
        if 'Статус товара' in df.columns and 'Дата создания' in df.columns:
            filtered_rows = []
            
            for idx, row in df.iterrows():
                days_left = calculate_days_until_expiration(
                    row['Дата создания'], 
                    row['Статус товара'], 
                    order_type
                )
                if days_left <= 2:
                    filtered_rows.append(idx)
            
            df = df.loc[filtered_rows]
            print(f"📊 После фильтрации по срокам: {len(df)} строк")
        
        # Применяем фильтрацию MX
        df = apply_mx_filters(df, processing_type)
        print(f"📊 После фильтрации MX: {len(df)} строк")
        
        # Добавляем столбец срока списания и сортируем
        if len(df) > 0:
            df = add_expiration_column(df, order_type)
            df = sort_by_priority(df)
        
        # Определяем какие столбцы оставить в финальном результате
        columns_to_keep = ['Дата создания', 'Гофра', 'Товар', 'Наименование', 'Стоимость', 'Осталось до списания']
        
        # Оставляем только нужные столбцы (если они есть в DataFrame)
        existing_columns = [col for col in columns_to_keep if col in df.columns]
        df = df[existing_columns]
        
        # Создаем файл в памяти
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        
        # Сохраняем файл для скачивания
        filename = f"обработанный_{secure_filename(file.filename)}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        with open(filepath, 'wb') as f:
            f.write(output.getvalue())
        
        # Генерируем сводку
        summary_text = generate_summary(df, order_type, processing_type)
        
        return jsonify({
            'success': True,
            'summary': summary_text,
            'download_url': f'/download/{filename}',
            'filename': filename
        })
        
    except Exception as e:
        logger.error(f"Ошибка при обработке файла: {e}")
        return jsonify({'success': False, 'error': f'Произошла ошибка при обработке: {str(e)}'})

@app.route('/download/<filename>')
def download_file(filename):
    """Скачивание обработанного файла"""
    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'success': False, 'error': f'Файл не найден: {str(e)}'})

if __name__ == '__main__':
    print("🚀 Запускаю веб-приложение...")
    print("📝 Откройте в браузере: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)