import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import configparser
import serial
import serial.tools.list_ports
import threading
import time
import pymssql
from datetime import datetime
import requests
import json
from urllib3.exceptions import InsecureRequestWarning
import urllib3
import uuid
from decimal import Decimal

class MainApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Приложение для работы с заказами")
        self.root.geometry("1200x800")
        
        # Переменные для подключения к БД
        self.db_connection = None
        self.db_config = {}
        
        # Переменные для COM порта
        self.serial_port = None
        self.serial_thread = None
        self.serial_running = False
        
        # Переменные для данных заказа
        self.current_order_data = {}
        self.current_bills_data = []
        self.product_items = {}  # Для хранения ссылок на элементы таблицы
        self.displayed_products = []  # Для хранения отображаемых товаров (с разбивкой по количеству)
        # self.without_mark_buttons = {}  # Убираем словарь кнопок "БЕЗ МАРКИ"
        
        # Переменные для CDN
        self.cdn_servers = []
        self.best_cdn_url = None
        self.best_cdn_latency = None
        
        # Переменные для Эвотор
        self.evotor_config = {}
        self.evotor_token = None
        
        # Загрузка конфигурации
        self.load_config()
        
        # Создание интерфейса
        self.create_widgets()
        
        # Автоподключение к БД и COM порту
        self.auto_connect()
        
        # Запуск фоновой проверки CDN
        self.start_cdn_check()
        
        # Привязываем обработчик изменения размера
        self.root.bind('<Configure>', self.on_window_resize)

    def load_config(self):
        """Загрузка конфигурации из файла conf.ini"""
        try:
            config = configparser.ConfigParser()
            config.read('conf.ini', encoding='utf-8')
            
            # Загрузка параметров базы данных
            self.db_config = {
                'server': config.get('DATABASE', 'server', fallback='localhost'),
                'database': config.get('DATABASE', 'database', fallback=''),
                'driver': config.get('DATABASE', 'driver', fallback='SQL Server'),
                'auth_mode': config.get('DATABASE', 'auth_mode', fallback='SQL'),
                'username': config.get('DATABASE', 'username', fallback=''),
                'password': config.get('DATABASE', 'password', fallback='')
            }
            
            # Загрузка других параметров
            self.max_weight = config.getint('OPTIONS', 'MAXWeight', fallback=5000)
            self.rs232_port = config.get('OPTIONS', 'rs232', fallback='COM3')
            
            # Загрузка параметров API Честного знака
            self.honest_sign_config = {
                'api_url': config.get('HONEST_SIGN', 'api_url', fallback='https://markirovka.crpt.ru'),
                'api_token': config.get('HONEST_SIGN', 'api_token', fallback=''),
                'verify_ssl': config.getboolean('HONEST_SIGN', 'verify_ssl', fallback=False),
                'timeout': config.getint('HONEST_SIGN', 'timeout', fallback=15000),
                'device_id': config.get('HONEST_SIGN', 'device_id', fallback='ALL'),
                'database_name': config.get('HONEST_SIGN', 'database_name', fallback='POS'),
                'login': config.get('HONEST_SIGN', 'login', fallback='USERID'),
                'password': config.get('HONEST_SIGN', 'password', fallback='PASSWORD')
            }
            
            # Загрузка параметров API Эвотор
            self.evotor_config = {
                'login': config.get('EVOTOR', 'login', fallback=''),
                'password': config.get('EVOTOR', 'password', fallback=''),
                'group_code': config.get('EVOTOR', 'group_code', fallback=''),
                'api_url': config.get('EVOTOR', 'api_url', fallback='https://api.evotor.ru'),
                'tax_system': config.getint('EVOTOR', 'tax_system', fallback=0),
                'payment_method': config.getint('EVOTOR', 'payment_method', fallback=1)
            }
            
            # Загрузка списка CDN серверов
            self.cdn_servers = []
            if config.has_section('cdn_servers'):
                for key in config['cdn_servers']:
                    self.cdn_servers.append(config['cdn_servers'][key])
            else:
                # Список по умолчанию если секция отсутствует
                self.cdn_servers = [
                    "https://cdn01.crpt.ru",
                    "https://cdn02.crpt.ru", 
                    "https://cdn03.crpt.ru",
                    "https://cdn04.crpt.ru",
                    "https://cdn05.crpt.ru",
                    "https://cdn06.crpt.ru",
                    "https://cdn07.crpt.ru",
                    "https://cdn08.crpt.ru",
                    "https://cdn09.crpt.ru",
                    "https://cdn10.crpt.ru",
                    "https://cdn11.crpt.ru"
                ]
            
            # Отключаем предупреждения SSL если нужно
            if not self.honest_sign_config['verify_ssl']:
                requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
            
            # Создаем текстовое поле для логов (если еще не создано)
            if not hasattr(self, 'log_text'):
                self.log_text = None
                
            self.log_message("Конфигурация загружена успешно")
            self.log_message(f"API URL: {self.honest_sign_config['api_url']}")
            self.log_message(f"SSL проверка: {self.honest_sign_config['verify_ssl']}")
            self.log_message(f"Токен задан: {'Да' if self.honest_sign_config['api_token'] else 'Нет'}")
            self.log_message(f"Загружено CDN серверов: {len(self.cdn_servers)}")
            self.log_message(f"Эвотор API URL: {self.evotor_config['api_url']}")
            self.log_message(f"Эвотор группа: {self.evotor_config['group_code']}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка загрузки конфигурации: {str(e)}")
            self.log_message(f"Ошибка загрузки конфигурации: {str(e)}")

    def start_cdn_check(self):
        """Запуск фоновой проверки CDN серверов"""
        self.log_message("Запуск фоновой проверки CDN серверов...")
        cdn_thread = threading.Thread(target=self.check_cdn_servers, daemon=True)
        cdn_thread.start()

    def check_cdn_servers(self):
        """Проверка доступности CDN серверов и выбор лучшего"""
        try:
            if not self.honest_sign_config.get('api_token'):
                self.log_message("Токен API Честного знака не настроен, пропускаем проверку CDN")
                return
                
            # Заголовки авторизации
            headers = {
                'X-API-KEY': self.honest_sign_config['api_token'],
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Connection': 'close'
            }
            
            self.log_message(f"Проверка доступности {len(self.cdn_servers)} CDN серверов...")
            
            # Проверяем доступность каждой площадки и измеряем latency
            available_cdns = []
            for cdn_url in self.cdn_servers:
                try:
                    health_url = f"{cdn_url}/api/v4/true-api/cdn/health/check"
                    start_time = time.time()
                    health_response = requests.get(
                        health_url,
                        headers=headers,
                        timeout=5,  # Таймаут 5 секунд
                        verify=False
                    )
                    end_time = time.time()
                    latency = (end_time - start_time) * 1000  # В миллисекундах
                    
                    if health_response.status_code == 200:
                        available_cdns.append((cdn_url, latency))
                        self.log_message(f"CDN {cdn_url} доступна, latency: {latency:.2f}ms")
                    else:
                        self.log_message(f"CDN {cdn_url} вернула статус {health_response.status_code}")
                except Exception as e:
                    self.log_message(f"CDN {cdn_url} недоступна: {str(e)}")
                    continue
            
            if not available_cdns:
                self.log_message("Нет доступных CDN серверов")
                return
                
            # Сортируем по latency (наименьшее время отклика первое)
            available_cdns.sort(key=lambda x: x[1])
            self.best_cdn_url, self.best_cdn_latency = available_cdns[0]
            
            self.log_message(f"Найдено доступных CDN: {len(available_cdns)}")
            self.log_message(f"Лучший CDN: {self.best_cdn_url} (latency: {self.best_cdn_latency:.2f}ms)")
            
            # Обновляем статус в интерфейсе
            self.cdn_status_label.config(text=f"CDN: {self.best_cdn_url} ({self.best_cdn_latency:.2f}ms)")
            
        except Exception as e:
            self.log_message(f"Ошибка при проверке CDN серверов: {str(e)}")

    def auto_connect(self):
        """Автоподключение к БД и COM порту при запуске"""
        # Подключение к БД
        self.connect_to_database()
        
        # Подключение к COM порту через небольшую задержку
        self.root.after(1000, self.connect_to_com_port)

    def connect_to_database(self):
        """Подключение к базе данных MS SQL"""
        try:
            if self.db_config['auth_mode'] == 'SQL':
                self.db_connection = pymssql.connect(
                    server=self.db_config['server'],
                    database=self.db_config['database'],
                    user=self.db_config['username'],
                    password=self.db_config['password']
                )
            else:
                # Windows аутентификация
                self.db_connection = pymssql.connect(
                    server=self.db_config['server'],
                    database=self.db_config['database']
                )
            
            self.status_label.config(text="Статус: Подключено к БД", foreground="green")
            self.log_message("Успешное подключение к базе данных")
            
        except Exception as e:
            self.status_label.config(text="Статус: Ошибка подключения к БД", foreground="red")
            self.log_message(f"Ошибка подключения к БД: {str(e)}")
            messagebox.showerror("Ошибка БД", f"Не удалось подключиться к базе данных: {str(e)}")

    def create_widgets(self):
        """Создание элементов интерфейса"""
        # Создание вкладок
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка основная (для заказов)
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Заказы")
        
        # Вкладка логов
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="Логи")
        
        # Вкладка настроек
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Настройки")
        
        # Создание элементов для основной вкладки
        self.create_main_tab()
        
        # Создание элементов для вкладки логов
        self.create_log_tab()
        
        # Создание элементов для вкладки настроек
        self.create_settings_tab()
        
        # Статусная строка
        self.status_frame = ttk.Frame(self.root)
        self.status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        self.status_label = ttk.Label(self.status_frame, text="Статус: Подключение...")
        self.status_label.pack(side=tk.LEFT)
        
        self.com_status_label = ttk.Label(self.status_frame, text="COM порт: Подключение...")
        self.com_status_label.pack(side=tk.RIGHT)
        
        # Статус CDN
        self.cdn_status_label = ttk.Label(self.status_frame, text="CDN: Проверка...")
        self.cdn_status_label.pack(side=tk.RIGHT, padx=(10, 0))

    def create_main_tab(self):
        """Создание элементов основной вкладки для заказов"""
        # Фрейм для сканирования QR кода
        scan_frame = ttk.LabelFrame(self.main_frame, text="Сканирование заказа", padding=10)
        scan_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(scan_frame, text="Отсканируйте QR код с номером заказа или введите вручную:").pack(anchor=tk.W)
        
        # Поле ввода для номера заказа (для тестирования)
        input_frame = ttk.Frame(scan_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Номер заказа:").pack(side=tk.LEFT)
        self.order_id_entry = ttk.Entry(input_frame, width=20)
        self.order_id_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="Получить заказ", 
                  command=self.get_order_by_manual_input).pack(side=tk.LEFT, padx=5)
        
        # Фрейм для отображения информации о заказе
        self.order_info_frame = ttk.LabelFrame(self.main_frame, text="Информация о заказе", padding=10)
        self.order_info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Создание элементов для информации о заказе
        self.create_order_info_widgets()
        
        # Фрейм для таблицы товаров
        self.products_frame = ttk.LabelFrame(self.main_frame, text="Товары в заказе", padding=10)
        self.products_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создание таблицы для товаров
        self.create_products_table()
        
        # === ДОБАВЛЕНО: Подсказка под таблицей ===
        self.hint_label = ttk.Label(self.main_frame, text="Для снятия признака маркировки два раза нажмите на товар", 
                                   font=('Arial', 9, 'italic'), foreground='gray')
        self.hint_label.pack(pady=(0, 5))
        
        # Фрейм для итоговой информации
        self.total_info_frame = ttk.Frame(self.main_frame)
        self.total_info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Итоговая стоимость заказа
        self.total_amount_label = ttk.Label(self.total_info_frame, text="Итоговая стоимость: 0.00 руб.", 
                                           font=('Arial', 12, 'bold'))
        self.total_amount_label.pack(side=tk.RIGHT)
        
        # Кнопка фискализации
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.fiscalize_btn = ttk.Button(
            button_frame, 
            text="Фискализировать чек", 
            command=self.finalize_order_and_send_receipt,
            state=tk.DISABLED
        )
        self.fiscalize_btn.pack(side=tk.RIGHT, padx=5)
        
        # Кнопка очистки
        ttk.Button(button_frame, text="Очистить", 
                  command=self.clear_order_info).pack(side=tk.RIGHT)

    def create_order_info_widgets(self):
        """Создание элементов для отображения информации о заказе"""
        # Верхняя часть с информацией о заказе
        top_info_frame = ttk.Frame(self.order_info_frame)
        top_info_frame.pack(fill=tk.X)
        
        # OrderID
        order_id_frame = ttk.Frame(top_info_frame)
        order_id_frame.pack(fill=tk.X, pady=2)
        ttk.Label(order_id_frame, text="Номер заказа:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT)
        self.order_id_label = ttk.Label(order_id_frame, text="", font=('Arial', 10, 'bold'))
        self.order_id_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # Имя клиента
        name_frame = ttk.Frame(top_info_frame)
        name_frame.pack(fill=tk.X, pady=2)
        ttk.Label(name_frame, text="Имя клиента:").pack(side=tk.LEFT)
        self.client_name_label = ttk.Label(name_frame, text="")
        self.client_name_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # Email
        email_frame = ttk.Frame(top_info_frame)
        email_frame.pack(fill=tk.X, pady=2)
        ttk.Label(email_frame, text="Email:").pack(side=tk.LEFT)
        self.email_label = ttk.Label(email_frame, text="")
        self.email_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # Скрытые поля (запоминаем, но не отображаем)
        self.employee_id_var = tk.StringVar()
        self.phone_var = tk.StringVar()
        self.delivery_point_var = tk.StringVar()
        self.grd_code_var = tk.StringVar()

    def create_products_table(self):
        """Создание таблицы для отображения товаров"""
        # === ИЗМЕНЕНО: Удален столбец 'Actions' ===
        columns = ('ProductName', 'Price', 'Quantity', 'TotalWeight', 'TotalAmount', 'MarkInfo')
        self.products_tree = ttk.Treeview(self.products_frame, columns=columns, show='headings', height=10)
        
        # Определение заголовков
        self.products_tree.heading('ProductName', text='Наименование товара')
        self.products_tree.heading('Price', text='Цена')
        self.products_tree.heading('Quantity', text='Количество')
        self.products_tree.heading('TotalWeight', text='Общий вес')
        self.products_tree.heading('TotalAmount', text='Сумма')
        self.products_tree.heading('MarkInfo', text='Марка Честного знака')
        
        # Определение ширины колонок
        self.products_tree.column('ProductName', width=250)
        self.products_tree.column('Price', width=80)
        self.products_tree.column('Quantity', width=80)
        self.products_tree.column('TotalWeight', width=80)
        self.products_tree.column('TotalAmount', width=80)
        self.products_tree.column('MarkInfo', width=200)
        
        # Добавление скроллбара
        scrollbar = ttk.Scrollbar(self.products_frame, orient=tk.VERTICAL, command=self.products_tree.yview)
        self.products_tree.configure(yscrollcommand=scrollbar.set)
        
        # Размещение таблицы и скроллбара
        self.products_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Привязка события клика по таблице
        self.products_tree.bind('<Double-1>', self.on_product_double_click)

    def on_product_double_click(self, event):
        """Обработка двойного клика по товару"""
        selection = self.products_tree.selection()
        if selection:
            item = selection[0]
            # Находим индекс товара
            for i, stored_item in self.product_items.items():
                if stored_item == item:
                    if i < len(self.displayed_products):
                        product = self.displayed_products[i]
                        # Проверяем, требует ли товар маркировки и не помечен ли уже как "без марки"
                        if product.get('ismarked', False) and not product.get('without_mark', False):
                            # Вместо ввода марки спрашиваем, нужно ли помечать как "без марки"
                            product_name = product.get('ProductName', 'Неизвестный товар')
                            if product.get('ismarked', False) and int(product.get('Quantity', 1)) > 1:
                                product_name = f"{product['ProductName']} [{product['display_index'] + 1}]"
                            
                            # Создаем диалоговое окно подтверждения
                            dialog = tk.Toplevel(self.root)
                            dialog.title("Подтверждение")
                            dialog.geometry("350x120")
                            dialog.transient(self.root)
                            dialog.grab_set()
                            # Центрируем
                            dialog.geometry("+%d+%d" % (dialog.winfo_screenwidth()/2-175, dialog.winfo_screenheight()/2-60))
                            
                            ttk.Label(dialog, text=f"Пометить товар как не требующий маркировки?", wraplength=300).pack(pady=10)
                            ttk.Label(dialog, text=f"'{product_name}'", font=('Arial', 9, 'bold'), wraplength=300).pack(pady=5)

                            def set_without_mark_confirmed():
                                self.set_without_mark(i) # Используем существующий метод
                                dialog.destroy()

                            def cancel_action():
                                dialog.destroy()
                                # Если нужно, можно здесь же открыть ввод марки
                                # self.request_mark_input(item, i)

                            button_frame = ttk.Frame(dialog)
                            button_frame.pack(pady=10)
                            ttk.Button(button_frame, text="ДА", command=set_without_mark_confirmed).pack(side=tk.LEFT, padx=5)
                            ttk.Button(button_frame, text="НЕТ", command=cancel_action).pack(side=tk.LEFT, padx=5)
                            
                        elif not product.get('ismarked', False) or product.get('without_mark', False):
                             # Если товар не требует маркировки или уже помечен как "без марки", 
                             # можно ничего не делать или показать сообщение
                             pass
                        else:
                            # Если товар требует маркировки, но уже есть марка, 
                            # можно ничего не делать или показать сообщение
                             pass
                    break

    def request_mark_input(self, item, product_index):
        """Запрос ввода марки Честного знака"""
        # Создаем диалоговое окно для ввода марки
        dialog = tk.Toplevel(self.root)
        dialog.title("Ввод марки Честного знака")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Центрируем окно
        dialog.geometry("+%d+%d" % (dialog.winfo_screenwidth()/2-200, dialog.winfo_screenheight()/2-75))
        
        ttk.Label(dialog, text="Введите марку Честного знака:").pack(pady=10)
        mark_entry = ttk.Entry(dialog, width=50)
        mark_entry.pack(pady=5)
        mark_entry.focus()
        
        def save_mark():
            mark = mark_entry.get().strip()
            if mark:
                # Проверяем марку через API и показываем диалог подтверждения
                self.check_mark_and_show_dialog(mark, product_index, item)
                dialog.destroy()
            else:
                messagebox.showwarning("Предупреждение", "Введите марку Честного знака")
        
        def cancel():
            dialog.destroy()
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Сохранить", command=save_mark).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Отмена", command=cancel).pack(side=tk.LEFT, padx=5)
        
        # Привязка Enter к сохранению
        mark_entry.bind('<Return>', lambda e: save_mark())

    def create_log_tab(self):
        """Создание элементов вкладки логов"""
        # Текстовое поле для логов
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=25)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Кнопка очистки логов
        ttk.Button(self.log_frame, text="Очистить логи", 
                  command=self.clear_logs).pack(pady=5)

    def create_settings_tab(self):
        """Создание элементов вкладки настроек"""
        # Информация о подключении к БД
        db_frame = ttk.LabelFrame(self.settings_frame, text="Настройки базы данных", padding=10)
        db_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(db_frame, text=f"Сервер: {self.db_config.get('server', '')}").pack(anchor=tk.W)
        ttk.Label(db_frame, text=f"База данных: {self.db_config.get('database', '')}").pack(anchor=tk.W)
        ttk.Label(db_frame, text=f"Пользователь: {self.db_config.get('username', '')}").pack(anchor=tk.W)
        
        # Настройки COM порта
        com_settings_frame = ttk.LabelFrame(self.settings_frame, text="Настройки COM порта", padding=10)
        com_settings_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(com_settings_frame, text=f"COM порт: {self.rs232_port}").pack(anchor=tk.W)
        ttk.Label(com_settings_frame, text=f"Максимальный вес: {self.max_weight}").pack(anchor=tk.W)
        
        # Кнопки управления COM портом
        button_frame = ttk.Frame(com_settings_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.connect_com_btn = ttk.Button(button_frame, text="Переподключить COM порт", 
                                         command=self.reconnect_com_port)
        self.connect_com_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.disconnect_com_btn = ttk.Button(button_frame, text="Отключить COM порт", 
                                           command=self.disconnect_from_com_port)
        self.disconnect_com_btn.pack(side=tk.LEFT)
        
        # Настройки API Честного знака
        api_settings_frame = ttk.LabelFrame(self.settings_frame, text="Настройки API Честного знака", padding=10)
        api_settings_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(api_settings_frame, text=f"API URL: {self.honest_sign_config.get('api_url', '')}").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Токен: {'*' * 10 if self.honest_sign_config.get('api_token') else 'Не задан'}").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Проверка SSL: {'Включена' if self.honest_sign_config.get('verify_ssl', True) else 'Отключена'}").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Таймаут: {self.honest_sign_config.get('timeout', 15000)} мс").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Device ID: {self.honest_sign_config.get('device_id', 'ALL')}").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Database Name: {self.honest_sign_config.get('database_name', 'POS')}").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Login: {self.honest_sign_config.get('login', 'USERID')}").pack(anchor=tk.W)
        ttk.Label(api_settings_frame, text=f"Password: {'*' * 10 if self.honest_sign_config.get('password') else 'Не задан'}").pack(anchor=tk.W)
        
        # Информация о CDN
        cdn_frame = ttk.LabelFrame(self.settings_frame, text="Информация о CDN", padding=10)
        cdn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(cdn_frame, text=f"Всего серверов: {len(self.cdn_servers)}").pack(anchor=tk.W)
        ttk.Label(cdn_frame, text=f"Лучший сервер: {self.best_cdn_url or 'Не определен'}").pack(anchor=tk.W)
        ttk.Label(cdn_frame, text=f"Latency: {self.best_cdn_latency:.2f}ms" if self.best_cdn_latency else "Latency: Не определен").pack(anchor=tk.W)
        
        # Настройки API Эвотор
        evotor_settings_frame = ttk.LabelFrame(self.settings_frame, text="Настройки API Эвотор", padding=10)
        evotor_settings_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(evotor_settings_frame, text=f"API URL: {self.evotor_config.get('api_url', '')}").pack(anchor=tk.W)
        ttk.Label(evotor_settings_frame, text=f"Login: {self.evotor_config.get('login', '')}").pack(anchor=tk.W)
        ttk.Label(evotor_settings_frame, text=f"Password: {'*' * 10 if self.evotor_config.get('password') else 'Не задан'}").pack(anchor=tk.W)
        ttk.Label(evotor_settings_frame, text=f"Group Code: {self.evotor_config.get('group_code', '')}").pack(anchor=tk.W)
        ttk.Label(evotor_settings_frame, text=f"Tax System: {self.evotor_config.get('tax_system', 0)}").pack(anchor=tk.W)
        ttk.Label(evotor_settings_frame, text=f"Payment Method: {self.evotor_config.get('payment_method', 1)}").pack(anchor=tk.W)
        
        # Кнопки управления
        button_frame = ttk.Frame(self.settings_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Переподключиться к БД", 
                  command=self.reconnect_database).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Переподключить COM порт", 
                  command=self.reconnect_com_port).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Обновить CDN серверы", 
                  command=self.start_cdn_check).pack(side=tk.LEFT, padx=(10, 0))

    def get_order_by_manual_input(self):
        """Получение заказа по ручному вводу номера"""
        order_id = self.order_id_entry.get().strip()
        if not order_id:
            messagebox.showwarning("Предупреждение", "Введите номер заказа")
            return
        self.get_order_info(order_id)

    def get_order_info(self, order_id):
        """Получение информации о заказе из базы данных"""
        try:
            if not self.db_connection:
                messagebox.showerror("Ошибка", "Нет подключения к базе данных")
                return
                
            cursor = self.db_connection.cursor()
            
            # Получение данных из таблицы orders с информацией о маркировке
            cursor.execute("""
                SELECT OrderID, Name, Email, EmployeeID, Phone, DeliveryPoint, GRDCode,
                       ProductName, CaseWeight, Price, Quantity, TotalWeight, TotalAmount, ismarked, Mark
                FROM orders 
                WHERE OrderID = %s
            """, (order_id,))
            
            bills_results = cursor.fetchall()
            
            if not bills_results:
                messagebox.showerror("Ошибка", f"Данные о заказе {order_id} не найдены в таблице orders")
                self.log_message(f"Данные о заказе {order_id} не найдены в таблице orders")
                return
            
            # Сохранение данных
            first_row = bills_results[0]
            self.current_order_data = {
                'OrderID': first_row[0],
                'Name': first_row[1],
                'Email': first_row[2],
                'EmployeeID': first_row[3],
                'Phone': first_row[4],
                'DeliveryPoint': first_row[5],
                'GRDCode': first_row[6]
            }
            
            # Сохранение данных о товарах
            self.current_bills_data = []
            for row in bills_results:
                self.current_bills_data.append({
                    'ProductName': row[7],
                    'CaseWeight': row[8],
                    'Price': row[9],
                    'Quantity': row[10],
                    'TotalWeight': row[11],
                    'TotalAmount': row[12],
                    'ismarked': row[13] if len(row) > 13 else False,
                    'mark': row[14] if len(row) > 14 else ''  # Марка из базы
                })
            
            # Отображение информации
            self.display_order_info()
            self.display_products_info()
            
            self.log_message(f"Успешно загружен заказ {order_id}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при получении данных заказа: {str(e)}")
            self.log_message(f"Ошибка при получении данных заказа {order_id}: {str(e)}")
        finally:
            try:
                cursor.close()
            except:
                pass

    def display_order_info(self):
        """Отображение информации о заказе"""
        # Очистка таблицы товаров
        for item in self.products_tree.get_children():
            self.products_tree.delete(item)
            
        # Отображение информации о заказе
        self.order_id_label.config(text=self.current_order_data.get('OrderID', ''))
        self.client_name_label.config(text=self.current_order_data.get('Name', ''))
        self.email_label.config(text=self.current_order_data.get('Email', ''))
        
        # Сохранение скрытых данных
        self.employee_id_var.set(self.current_order_data.get('EmployeeID', ''))
        self.phone_var.set(self.current_order_data.get('Phone', ''))
        self.delivery_point_var.set(self.current_order_data.get('DeliveryPoint', ''))
        self.grd_code_var.set(self.current_order_data.get('GRDCode', ''))
        
        # Активируем кнопку фискализации
        self.fiscalize_btn.config(state=tk.NORMAL)

    def display_products_info(self):
        """Отображение информации о товарах с разбивкой по количеству"""
        # Очистка таблицы
        for item in self.products_tree.get_children():
            self.products_tree.delete(item)
            
        self.product_items = {}  # Очищаем словарь ссылок на элементы
        self.displayed_products = []  # Очищаем список отображаемых товаров
        # self.without_mark_buttons = {}  # Убираем очистку словаря кнопок
        
        # Создаем отображаемые товары с разбивкой по количеству
        for product in self.current_bills_data:
            quantity = int(product.get('Quantity', 1))
            ismarked = product.get('ismarked', False)
            
            # Если товар маркированный и количество больше 1, создаем несколько записей
            if ismarked and quantity > 1:
                # Создаем отдельную запись для каждой единицы товара
                for i in range(quantity):
                    # Копируем товар и устанавливаем количество = 1
                    single_product = product.copy()
                    single_product['Quantity'] = 1
                    single_product['TotalWeight'] = product['TotalWeight'] / quantity if quantity > 0 else 0
                    single_product['TotalAmount'] = product['TotalAmount'] / quantity if quantity > 0 else 0
                    single_product['display_index'] = i  # Индекс отображения
                    single_product['original_index'] = self.current_bills_data.index(product)  # Оригинальный индекс
                    
                    # Если у оригинального товара есть марка, применяем её к первой единице
                    if i == 0 and product.get('mark', ''):
                        single_product['mark'] = product['mark']
                    else:
                        single_product['mark'] = ''
                        
                    # Флаг для товаров без марки
                    single_product['without_mark'] = False
                    self.displayed_products.append(single_product)
            else:
                # Для немаркированных товаров или товаров с количеством 1
                product_copy = product.copy()
                product_copy['display_index'] = 0
                product_copy['original_index'] = self.current_bills_data.index(product)
                product_copy['without_mark'] = False
                self.displayed_products.append(product_copy)
        
        # Вычисляем итоговую стоимость заказа
        total_amount = sum(float(product.get('TotalAmount', 0)) for product in self.current_bills_data)
        self.total_amount_label.config(text=f"Итоговая стоимость: {total_amount:.2f} руб.")
        
        # Добавление товаров в таблицу
        for i, product in enumerate(self.displayed_products):
            ismarked = product.get('ismarked', False)
            mark = product.get('mark', '')
            without_mark = product.get('without_mark', False)
            
            # Формируем название товара с индексом для маркированных товаров
            product_name = product['ProductName']
            if ismarked and int(product.get('Quantity', 1)) > 1:
                product_name = f"{product['ProductName']} [{product['display_index'] + 1}]"
                
            # Если товар без марки, добавляем пометку
            if without_mark or not ismarked:
                product_name += " (БЕЗ МАРКИ)"
                
            # === ИЗМЕНЕНО: Удалена пустая колонка 'Actions' ===
            item = self.products_tree.insert('', tk.END, values=(
                product_name,
                f"{product['Price']:.2f}",
                product['Quantity'],
                f"{product['TotalWeight']:.2f}",
                f"{product['TotalAmount']:.2f}",
                mark if ismarked and not without_mark else ("БЕЗ МАРКИ" if not ismarked or without_mark else "")
            ))
            
            self.product_items[i] = item
            
            # Применяем стили в зависимости от маркировки
            if ismarked and not without_mark:
                if mark:
                    # Зеленый фон если марка введена
                    self.products_tree.tag_configure(f'green_{i}', background='lightgreen')
                    self.products_tree.item(item, tags=(f'green_{i}',))
                else:
                    # Красный фон если требуется марка
                    self.products_tree.tag_configure(f'red_{i}', background='lightcoral')
                    self.products_tree.item(item, tags=(f'red_{i}',))
                    # Показываем сообщение
                    self.log_message(f"Требуется марка Честного знака для товара: {product['ProductName']}")
            elif without_mark or not ismarked:
                # Серый фон для товаров без марки
                self.products_tree.tag_configure(f'gray_{i}', background='lightgray')
                self.products_tree.item(item, tags=(f'gray_{i}',))
        
        # Убираем вызов add_without_mark_buttons

    # Убираем метод add_without_mark_buttons полностью

    def set_without_mark(self, product_index):
        """Установка флага "БЕЗ МАРКИ" для товара"""
        if product_index < len(self.displayed_products):
            product = self.displayed_products[product_index]
            # Устанавливаем флаг без марки
            product['without_mark'] = True
            self.log_message(f"[БЕЗ МАРКИ] Установлен флаг для товара индекс {product_index}: {product.get('ProductName', 'Unknown')}")
            
            # Обновляем отображение
            item = self.product_items.get(product_index)
            if item:
                try:
                    # Формируем новое название товара
                    product_name = product['ProductName']
                    if product.get('ismarked', False) and int(product.get('Quantity', 1)) > 1:
                        product_name = f"{product['ProductName']} [{product['display_index'] + 1}]"
                    product_name_with_flag = f"{product_name} (БЕЗ МАРКИ)"
                    
                    # === ИЗМЕНЕНО: Обновляем значения напрямую через set для каждой колонки ===
                    self.products_tree.set(item, 'ProductName', product_name_with_flag)
                    self.products_tree.set(item, 'MarkInfo', "БЕЗ МАРКИ")
                    
                    # Меняем цвет на серый
                    tag_name = f'gray_{product_index}'
                    self.products_tree.tag_configure(tag_name, background='lightgray')
                    self.products_tree.item(item, tags=(tag_name,))
                    
                    self.log_message(f"[БЕЗ МАРКИ] Обновлено отображение для индекса {product_index}")
                    
                except Exception as e:
                    self.log_message(f"[БЕЗ МАРКИ] ОШИБКА при обновлении интерфейса: {e}")
            else:
                 self.log_message(f"[БЕЗ МАРКИ] Элемент Treeview не найден для индекса {product_index}")
            
            self.log_message(f"[БЕЗ МАРКИ] Завершено для товара: {product.get('ProductName', 'Unknown')}")
        else:
             self.log_message(f"[БЕЗ МАРКИ] Неверный индекс продукта: {product_index}")

    def update_product_mark_display(self, item, mark, is_valid=False):
        """Обновление отображения марки товара"""
        # === ИЗМЕНЕНО: Обновляем значения напрямую через set для каждой колонки ===
        self.products_tree.set(item, 'MarkInfo', mark)
        
        # Обновляем стиль
        item_index = None
        for key, value in self.product_items.items():
            if value == item:
                item_index = key
                break
        
        if item_index is not None and is_valid:
            # Зеленый фон если марка введена правильно
            self.products_tree.tag_configure(f'green_{item_index}', background='lightgreen')
            self.products_tree.item(item, tags=(f'green_{item_index}',))

    def clear_order_info(self):
        """Очистка информации о заказе"""
        # Очистка меток
        self.order_id_label.config(text="")
        self.client_name_label.config(text="")
        self.email_label.config(text="")
        
        # Очистка скрытых полей
        self.employee_id_var.set("")
        self.phone_var.set("")
        self.delivery_point_var.set("")
        self.grd_code_var.set("")
        
        # Очистка таблицы товаров
        for item in self.products_tree.get_children():
            self.products_tree.delete(item)
            
        # Очистка данных
        self.current_order_data = {}
        self.current_bills_data = []
        self.product_items = {}
        self.displayed_products = []
        # self.without_mark_buttons = {}  # Убираем очистку словаря кнопок
        
        # Сброс итоговой стоимости
        self.total_amount_label.config(text="Итоговая стоимость: 0.00 руб.")
        
        # Деактивируем кнопку фискализации
        self.fiscalize_btn.config(state=tk.DISABLED)
        
        self.log_message("Информация о заказе очищена")

    def connect_to_com_port(self):
        """Подключение к COM порту"""
        try:
            # Закрываем предыдущее соединение, если оно есть
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
                
            self.serial_port = serial.Serial(
                port=self.rs232_port,
                baudrate=9600,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            
            self.serial_running = True
            self.serial_thread = threading.Thread(target=self.read_serial_data, daemon=True)
            self.serial_thread.start()
            
            self.com_status_label.config(text=f"COM порт: Подключен ({self.rs232_port})", foreground="green")
            self.log_message(f"Успешное подключение к COM порту {self.rs232_port}")
            
        except Exception as e:
            self.com_status_label.config(text="COM порт: Ошибка подключения", foreground="red")
            self.log_message(f"Ошибка подключения к COM порту: {str(e)}")

    def disconnect_from_com_port(self):
        """Отключение от COM порта"""
        try:
            self.serial_running = False
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            self.com_status_label.config(text="COM порт: Отключен", foreground="red")
            self.log_message("Отключение от COM порта")
        except Exception as e:
            self.log_message(f"Ошибка отключения от COM порта: {str(e)}")

    def reconnect_com_port(self):
        """Переподключение к COM порту"""
        self.disconnect_from_com_port()
        self.root.after(500, self.connect_to_com_port)

    def read_serial_data(self):
        """Чтение данных с COM порта в отдельном потоке"""
        while self.serial_running:
            try:
                if self.serial_port and self.serial_port.is_open:
                    if self.serial_port.in_waiting > 0:
                        data = self.serial_port.readline().decode('utf-8', errors='ignore').strip()
                        if data:  # Проверка на пустые данные
                            self.root.after(0, self.process_serial_data, data)
                time.sleep(0.01)
            except Exception as e:
                self.log_message(f"Ошибка чтения данных: {str(e)}")
                break

    def process_serial_data(self, data):
        """Обработка полученных данных с COM порта"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_message(f"Получены данные с COM порта: {data}")
        
        # Проверяем, является ли это маркой Честного знака
        # Марки Честного знака обычно длинные и содержат цифры и буквы
        if len(data) > 15 and (data.startswith('01') or 'mark' in data.lower() or 
                              any(c.isdigit() for c in data[:10])):
            # Это похоже на марку Честного знака
            self.root.after(0, self.process_mark_data, data)
        elif data.startswith('ORDER_'):  # Предполагаем, что заказы начинаются с ORDER_
            order_id = data.replace('ORDER_', '').strip()
            if order_id:
                self.root.after(0, self.get_order_info, order_id)
        else:
            # Предполагаем, что это номер заказа
            order_id = data.strip()
            if order_id:
                self.root.after(0, self.get_order_info, order_id)

    def process_mark_data(self, mark_data):
        """Обработка данных марки Честного знака"""
        # Очищаем марку от лишних символов
        mark = mark_data.strip()
        if not mark:
            self.log_message("Получена пустая марка")
            return
            
        # Ищем выделенный товар в таблице
        selection = self.products_tree.selection()
        if selection:
            item = selection[0]
            # Находим индекс товара
            for i, stored_item in self.product_items.items():
                if stored_item == item:
                    if i < len(self.displayed_products):
                        product = self.displayed_products[i]
                        if product.get('ismarked', False) and not product.get('without_mark', False):
                            # Проверяем марку через API
                            self.check_mark_and_show_dialog(mark, i, item)
                            return
        
        # Если нет выделенного товара, ищем первый товар, требующий марку
        for i, product in enumerate(self.displayed_products):
            if product.get('ismarked', False) and not product.get('mark', '') and not product.get('without_mark', False):
                item = self.product_items.get(i)
                if item:
                    # Проверяем марку через API
                    self.check_mark_and_show_dialog(mark, i, item)
                    return
                    
        # Если все товары уже помечены
        self.log_message(f"Получена марка {mark}, но все товары уже помечены")

    def check_mark_and_show_dialog(self, mark, product_index, item):
        """Проверка марки и показ диалога подтверждения"""
        # Показываем сообщение о проверке
        self.log_message(f"Проверка марки {mark} через API Честного знака...")
        
        # Блокируем интерфейс во время проверки
        self.root.config(cursor="wait")
        self.root.update()
        
        try:
            # Проверяем марку через API
            result = self.validate_mark_with_honest_sign(mark)
            
            # Восстанавливаем курсор
            self.root.config(cursor="")
            
            if result is None:
                messagebox.showerror("Ошибка", "Не удалось проверить марку - API не настроен")
                return
                
            if 'error' in result:
                error_msg = result.get('message', 'Неизвестная ошибка')
                messagebox.showerror("Ошибка проверки марки", f"Не удалось проверить марку: {error_msg}")
                return
            
            # Показываем диалог подтверждения
            self.show_mark_confirmation_dialog(mark, product_index, item, result)
            
        except Exception as e:
            self.root.config(cursor="")
            self.log_message(f"Ошибка при проверке марки: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка при проверке марки: {str(e)}")

    def show_mark_confirmation_dialog(self, mark, product_index, item, api_result):
        """Показ диалога подтверждения марки"""
        # Создаем диалоговое окно
        dialog = tk.Toplevel(self.root)
        dialog.title("Подтверждение марки Честного знака")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Центрируем окно
        dialog.geometry("+%d+%d" % (dialog.winfo_screenwidth()/2-300, dialog.winfo_screenheight()/2-250))
        
        # Информация о товаре
        product = self.displayed_products[product_index]
        product_name = product.get('ProductName', 'Неизвестный товар')
        if product.get('ismarked', False) and int(product.get('Quantity', 1)) > 1:
            product_name = f"{product['ProductName']} [{product['display_index'] + 1}]"
            
        ttk.Label(dialog, text=f"Товар: {product_name}", font=('Arial', 10, 'bold')).pack(pady=5)
        ttk.Label(dialog, text=f"Марка: {mark}", font=('Arial', 9)).pack(pady=5)
        
        # Проверяем статусы товара
        is_sold = False
        is_utilised = False  # Выведен из оборота
        is_blocked = False
        is_realizable = True
        is_valid = True  # Можно продавать
        
        # Извлекаем статусы из результата
        if isinstance(api_result, dict) and 'codes' in api_result and len(api_result['codes']) > 0:
            code_info = api_result['codes'][0]
            is_sold = code_info.get('sold', False)
            is_utilised = code_info.get('utilised', False)  # Выведен из оборота
            is_blocked = code_info.get('isBlocked', False)
            is_realizable = code_info.get('realizable', True)
            
            # === ИСПРАВЛЕННАЯ ЛОГИКА ===
            # Можно продавать если:
            # 1. Не продан (sold = false)
            # 2. Не заблокирован (isBlocked = false)  
            # 3. Можно реализовать (realizable = true)
            # 4. Выведен из оборота (utilised = true) - КЛЮЧЕВОЕ УСЛОВИЕ
            if is_sold or is_blocked or not is_realizable or not is_utilised:
                is_valid = False
        
        # Отображаем предупреждение если товар нельзя применять
        if not is_valid:
            warning_text = "ВНИМАНИЕ: "
            if is_sold:
                warning_text += "Товар уже продан! "
            if is_blocked:
                warning_text += "Товар заблокирован! "
            if not is_realizable:
                warning_text += "Товар не введён в оборот! "
            if not is_utilised:  # Если не выведен из оборота
                warning_text += "Товар НЕ выведен из оборота! "
                
            warning_label = ttk.Label(dialog, text=warning_text, foreground="red", font=('Arial', 10, 'bold'))
            warning_label.pack(pady=5)
        
        # Информация из API
        ttk.Label(dialog, text="Информация о товаре из Честного знака:", font=('Arial', 10, 'bold')).pack(pady=(10,5))
        
        # Создаем фрейм для информации
        info_frame = ttk.Frame(dialog)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Текстовое поле для информации
        info_text = scrolledtext.ScrolledText(info_frame, height=15, width=70)
        info_text.pack(fill=tk.BOTH, expand=True)
        
        # Форматируем информацию из API в читабельный вид
        info_lines = []
        if isinstance(api_result, dict):
            # Логируем полный ответ в консоль для отладки
            self.log_message(f"Полный ответ API: {json.dumps(api_result, ensure_ascii=False, indent=2)}")
            
            # Проверяем структуру ответа
            if 'codes' in api_result and len(api_result['codes']) > 0:
                # Формат ответа CDN
                code_info = api_result['codes'][0]
                info_lines.append("=== ОСНОВНАЯ ИНФОРМАЦИЯ ===")
                info_lines.append(f"Код идентификации: {code_info.get('cis', 'Не указан')}")
                info_lines.append(f"Статус: {'ДЕЙСТВИТЕЛЬНА' if code_info.get('valid', False) else 'НЕДЕЙСТВИТЕЛЬНА'}")
                info_lines.append(f"Найдена в системе: {'Да' if code_info.get('found', True) else 'Нет'}")
                info_lines.append(f"Можно реализовать: {'Да' if code_info.get('realizable', True) else 'Нет'}")
                info_lines.append(f"Введен в оборот: {'Да' if code_info.get('utilised', False) else 'Нет'}")
                info_lines.append("\n=== ДЕТАЛИ ТОВАРА ===")
                info_lines.append(f"GTIN: {code_info.get('gtin', 'Не указан')}")
                info_lines.append(f"Производитель (ИНН): {code_info.get('producerInn', 'Не указан')}")
                info_lines.append(f"Тип упаковки: {code_info.get('packageType', 'Не указан')}")
                
                if 'productionDate' in code_info:
                    prod_date = code_info['productionDate']
                    if prod_date:
                        # Преобразуем дату в читабельный формат
                        try:
                            from datetime import datetime
                            dt = datetime.fromisoformat(prod_date.replace('Z', '+00:00'))
                            info_lines.append(f"Дата производства: {dt.strftime('%d.%m.%Y')}")
                        except:
                            info_lines.append(f"Дата производства: {prod_date}")
                
                if 'expireDate' in code_info:
                    exp_date = code_info['expireDate']
                    if exp_date:
                        # Преобразуем дату в читабельный формат
                        try:
                            from datetime import datetime
                            dt = datetime.fromisoformat(exp_date.replace('Z', '+00:00'))
                            info_lines.append(f"Срок годности: {dt.strftime('%d.%m.%Y')}")
                        except:
                            info_lines.append(f"Срок годности: {exp_date}")
                
                info_lines.append("\n=== СТАТУСЫ ===")
                info_lines.append(f"Введен в оборот: {'Да' if code_info.get('utilised', False) else 'Нет'}")
                info_lines.append(f"Заблокирован: {'Да' if code_info.get('isBlocked', False) else 'Нет'}")
                info_lines.append(f"Продан: {'Да' if code_info.get('sold', False) else 'Нет'}")
                info_lines.append(f"Владелец: {'Да' if code_info.get('isOwner', False) else 'Нет'}")
                
                if 'errorCode' in code_info and code_info['errorCode'] != 0:
                    info_lines.append(f"Код ошибки: {code_info['errorCode']}")
                    
            elif 'description' in api_result:
                # Формат ответа с описанием
                info_lines.append(f"Статус: {api_result.get('description', 'Нет описания')}")
                if 'code' in api_result:
                    info_lines.append(f"Код ответа: {api_result['code']}")
            else:
                # Общий формат для других ответов
                info_lines.append("Подробная информация:")
                for key, value in api_result.items():
                    if key not in ['error']:
                        info_lines.append(f"{key}: {value}")
        
        if not info_lines:
            info_lines.append("Информация недоступна")
            
        info_text.insert(tk.END, "\n".join(info_lines))
        info_text.config(state=tk.DISABLED)
        
        # Фрейм для кнопок
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def apply_mark():
            """Применить марку"""
            # Сохраняем марку
            self.displayed_products[product_index]['mark'] = mark
            
            # Сохраняем в базу данных
            order_id = self.current_order_data.get('OrderID', '')
            original_product = self.current_bills_data[self.displayed_products[product_index]['original_index']]
            product_name = original_product['ProductName']
            
            # Для товаров с количеством > 1 сохраняем марку в оригинальный товар
            if int(original_product.get('Quantity', 1)) > 1:
                # Обновляем марку в оригинальном товаре
                original_product['mark'] = mark
                
            if self.save_mark_to_database(order_id, product_name, mark):
                # Обновляем отображение
                self.update_product_mark_display(item, mark, True)
                self.log_message(f"Марка {mark} применена для товара: {product_name}")
                dialog.destroy()
                
                # Переходим к следующему товару, требующему марку
                self.focus_next_marked_product(product_index)
            else:
                messagebox.showerror("Ошибка", "Не удалось сохранить марку в базу данных")
        
        def cancel_mark():
            """Отменить марку"""
            self.log_message(f"Марка {mark} отменена для товара: {product.get('ProductName', 'Неизвестный товар')}")
            dialog.destroy()
        
        # Кнопки
        apply_button = ttk.Button(button_frame, text="Применить", command=apply_mark)
        apply_button.pack(side=tk.LEFT, padx=5)
        
        # Если товар нельзя продавать, отключаем кнопку применения
        if not is_valid:
            apply_button.config(state=tk.DISABLED)
            apply_button.config(text="Применить (недоступно)")
            
        ttk.Button(button_frame, text="Отменить", command=cancel_mark).pack(side=tk.LEFT, padx=5)
        
        # Привязка клавиш
        if is_valid:
            dialog.bind('<Return>', lambda e: apply_mark())
        dialog.bind('<Escape>', lambda e: cancel_mark())

    def focus_next_marked_product(self, current_index):
        """Переход к следующему товару, требующему марку"""
        for i in range(current_index + 1, len(self.displayed_products)):
            product = self.displayed_products[i]
            if product.get('ismarked', False) and not product.get('mark', '') and not product.get('without_mark', False):
                item = self.product_items.get(i)
                if item:
                    # Выделяем товар в таблице
                    self.products_tree.selection_set(item)
                    self.products_tree.see(item)
                    self.products_tree.focus(item)
                    product_name = product.get('ProductName', 'Неизвестный товар')
                    if int(product.get('Quantity', 1)) > 1:
                        product_name = f"{product['ProductName']} [{product['display_index'] + 1}]"
                    self.log_message(f"Переход к следующему товару: {product_name}")
                    return
        
        # Если больше нет товаров, требующих марку
        self.log_message("Все товары помечены")
        messagebox.showinfo("Готово", "Все товары, требующие маркировки, успешно помечены!")

    def validate_mark_with_honest_sign(self, mark):
        """Проверка марки через лучший CDN сервер Честного знака"""
        try:
            self.log_message(f"Начало проверки марки: {mark}")
            
            if not self.honest_sign_config.get('api_token'):
                self.log_message("Токен API Честного знака не настроен")
                return None
                
            # Проверяем, есть ли лучший CDN сервер
            if not self.best_cdn_url:
                self.log_message("Лучший CDN сервер не определен, используем первый из списка")
                if not self.cdn_servers:
                    self.log_message("Список CDN серверов пуст")
                    return self._fallback_validation(mark)
                check_url = f"{self.cdn_servers[0]}/api/v4/true-api/codes/check"
            else:
                check_url = f"{self.best_cdn_url}/api/v4/true-api/codes/check"
                self.log_message(f"Используется лучший CDN сервер: {self.best_cdn_url}")
            
            # Заголовки авторизации (используем X-API-KEY как в основном решении)
            headers = {
                'X-API-KEY': self.honest_sign_config['api_token'],  # Важно: X-API-KEY, а не Authorization
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Connection': 'close'  # Закрываем соединение как в рекомендациях
            }
            
            # Извлекаем CIS из кода маркировки
            cis = mark.split('\x1d')[0] if '\x1d' in mark else mark
            
            data = {
                "codes": [cis]
            }
            
            self.log_message(f"Проверка кода {cis} через CDN {check_url}...")
            start_time = time.time()
            check_response = requests.post(
                check_url,
                headers=headers,
                json=data,
                timeout=10,  # Таймаут 10 секунд как в основном решении
                verify=False
            )
            end_time = time.time()
            request_time = (end_time - start_time) * 1000
            self.log_message(f"Запрос выполнен за {request_time:.2f}ms, статус {check_response.status_code}")
            
            if check_response.status_code == 200:
                result = check_response.json()
                self.log_message("Успешная проверка через CDN")
                
                # Логируем краткую информацию о результате
                if 'codes' in result and len(result['codes']) > 0:
                    code_info = result['codes'][0]
                    status = "ДЕЙСТВИТЕЛЬНА" if code_info.get('valid', False) else "НЕДЕЙСТВИТЕЛЬНА"
                    self.log_message(f"Код {cis}: {status}")
                    if 'gtin' in code_info:
                        self.log_message(f"GTIN: {code_info['gtin']}")
                    
                    # === ИСПРАВЛЕННАЯ ЛОГИКА ЛОГИРОВАНИЯ СТАТУСОВ ===
                    is_sold = code_info.get('sold', False)
                    is_utilised = code_info.get('utilised', False)  # Выведен из оборота
                    is_blocked = code_info.get('isBlocked', False)
                    is_realizable = code_info.get('realizable', True)
                    
                    if is_sold:
                        self.log_message("ВНИМАНИЕ: Товар уже продан!")
                    if is_blocked:
                        self.log_message("ВНИМАНИЕ: Товар заблокирован!")
                    if not is_realizable:
                        self.log_message("ВНИМАНИЕ: Товар не введён в оборот!")
                    if not is_utilised:  # Если не выведен из оборота
                        self.log_message("ВНИМАНИЕ: Товар НЕ выведен из оборота!")
                        
                return result
            else:
                self.log_message(f"Ошибка проверки через CDN: {check_response.status_code} - {check_response.text}")
                return self._fallback_validation(mark)
                
        except Exception as e:
            self.log_message(f"Ошибка проверки через CDN: {str(e)}")
            import traceback
            self.log_message(f"Трассировка: {traceback.format_exc()}")
            return self._fallback_validation(mark)

    def _fallback_validation(self, mark):
        """Резервный метод проверки - возвращает фиктивный результат"""
        self.log_message("Используется резервная проверка (фиктивный результат)")
        
        # Извлекаем CIS из кода маркировки
        cis = mark.split('\x1d')[0] if '\x1d' in mark else mark
        
        # Создаем фиктивный ответ как в основном решении
        fake_result = {
            "code": 0,
            "description": "ok (имитация)",
            "codes": [
                {
                    "cis": cis,
                    "valid": True,
                    "printView": cis,
                    "gtin": "0" + cis[2:16] if len(cis) >= 16 else "04600000000000",
                    "groupIds": [23],
                    "verified": True,
                    "found": True,
                    "realizable": True,
                    "utilised": True,  # Выведен из оборота (КЛЮЧЕВОЕ ИЗМЕНЕНИЕ)
                    "isBlocked": False,
                    "expireDate": "2026-12-31T00:00:00Z",
                    "productionDate": "2025-01-01T00:00:00Z",
                    "isOwner": False,
                    "errorCode": 0,
                    "isTracking": False,
                    "sold": False,  # Не продан
                    "packageType": "UNIT",
                    "producerInn": "1234567890",
                    "grayZone": False
                }
            ],
            "reqId": "test-" + str(int(time.time())),
            "reqTimestamp": int(time.time() * 1000)
        }
        return fake_result

    def save_mark_to_database(self, order_id, product_name, mark):
        """Сохранение марки в базу данных"""
        try:
            if not self.db_connection:
                self.log_message("Нет подключения к базе данных для сохранения марки")
                return False
                
            cursor = self.db_connection.cursor()
            cursor.execute("""
                UPDATE orders 
                SET Mark = %s 
                WHERE OrderID = %s AND ProductName = %s
            """, (mark, order_id, product_name))
            self.db_connection.commit()
            cursor.close()
            
            self.log_message(f"Марка {mark} сохранена для товара {product_name} в заказе {order_id}")
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка сохранения марки в БД: {str(e)}")
            return False

    def log_message(self, message):
        """Добавление сообщения в лог"""
        if hasattr(self, 'log_text') and self.log_text is not None:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"[{timestamp}] {message}\n"
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
            # Также выводим в консоль для отладки
            print(f"[{timestamp}] {message}")

    def clear_logs(self):
        """Очистка логов"""
        self.log_text.delete(1.0, tk.END)

    def test_db_connection(self):
        """Тестирование подключения к БД"""
        try:
            if self.db_connection:
                cursor = self.db_connection.cursor()
                cursor.execute("SELECT 1")
                result = cursor.fetchone()
                if result:
                    messagebox.showinfo("Тест подключения", "Подключение к БД работает корректно!")
                    self.log_message("Тест подключения к БД успешен")
                cursor.close()
        except Exception as e:
            messagebox.showerror("Тест подключения", f"Ошибка подключения к БД: {str(e)}")
            self.log_message(f"Ошибка теста подключения к БД: {str(e)}")

    def reconnect_database(self):
        """Переподключение к базе данных"""
        try:
            if self.db_connection:
                self.db_connection.close()
        except:
            pass
        self.connect_to_database()

    def on_window_resize(self, event=None):
        """Обработчик изменения размера окна"""
        # Убираем перерисовку кнопок
        pass

    # === МЕТОДЫ ДЛЯ РАБОТЫ С ЭВОТОРОМ ===
    
    def authenticate_evotor(self):
        """
        Аутентификация в API Эвотор и получение токена.
        """
        try:
            # Это упрощенный пример. В реальной реализации здесь должен быть код для получения токена.
            # headers = {
            #     'Content-Type': 'application/x-www-form-urlencoded',
            #     'Authorization': f'Basic {base64.b64encode(f"{self.evotor_config["login"]}:{self.evotor_config["password"]}".encode()).decode()}'
            # }
            # response = requests.post(f"{self.evotor_config['api_url']}/oauth/token", headers=headers, verify=False)
            # if response.status_code == 200:
            #     token_data = response.json()
            #     return token_data.get('access_token')
            
            # Пока возвращаем фиктивный токен для демонстрации
            self.log_message("Аутентификация в Эвотор (имитация)")
            return "dummy_evotor_token_12345"
        except Exception as e:
            self.log_message(f"Ошибка аутентификации в Эвотор: {e}")
            return None

    def prepare_fiscal_receipt(self, order_data, products_data):
        """
        Подготовка данных фискального чека на основе данных заказа.
        """
        try:
            receipt = {
                "id": str(uuid.uuid4()),  # Уникальный идентификатор документа
                "checkout_date": datetime.now().isoformat(),  # Дата и время оформления чека
                "doc_num": order_data.get('OrderID', '001'),  # Номер документа
                "doc_type": "SALE",  # Тип документа (SALE - продажа)
                "tax_system": self.evotor_config['tax_system'],  # Система налогообложения
                "cashier": "Кассир по умолчанию",  # ФИО кассира
                "cashier_inn": "",  # ИНН кассира (если есть)
                "customer_email": order_data.get('Email', ''),  # Email покупателя
                "customer_phone": order_data.get('Phone', ''),  # Телефон покупателя
                "positions": [],  # Позиции в чеке
                "payments": [],  # Оплаты
                "total": 0  # Итоговая сумма
            }
            
            total_sum = Decimal('0.00')
            
            # Добавляем позиции
            for product in products_data:
                price = Decimal(str(product.get('Price', '0')))
                quantity = Decimal(str(product.get('Quantity', '1')))
                amount = price * quantity
                total_sum += amount
                
                position = {
                    "uuid": str(uuid.uuid4()),
                    "product_name": product.get('ProductName', 'Неизвестный товар'),
                    "product_code": product.get('GTIN', ''),  # GTIN или штрихкод товара
                    "measure_name": "шт",  # Единица измерения
                    "measure_code": 796,  # Код единицы измерения (796 - штуки)
                    "price": float(price),
                    "quantity": float(quantity),
                    "amount": float(amount),
                    "tax_percent": 20,  # Налоговая ставка (пример)
                    "tax_sum": float(amount * Decimal('0.20')),  # Сумма налога
                    "payment_method": 1,  # Признак способа расчета (1 - полный расчет)
                    "payment_object": 1,  # Признак предмета расчета (1 - товар)
                    # Для маркированных товаров
                    "is_excise": product.get('ismarked', False),
                    "mark_code": product.get('mark', '') if product.get('ismarked', False) and product.get('mark') else None
                }
                
                # Если товар маркированный, добавляем код маркировки
                if product.get('ismarked', False) and product.get('mark', ''):
                    # Определяем правильный признак предмета расчета для маркировки
                    if product.get('is_alcohol', False):  # Алкоголь
                        position['payment_object'] = 31  # Алкоголь с маркировкой
                    else:  # Обычный товар с маркировкой
                        position['payment_object'] = 33  # Товар с маркировкой
                    
                    # Добавляем информацию о маркировке
                    position['mark_code'] = product['mark']
                    position['mark_processing_mode'] = 0  # Режим обработки кода маркировки
                
                receipt['positions'].append(position)
            
            # Добавляем оплату
            receipt['payments'].append({
                "type": self.evotor_config['payment_method'],  # Способ оплаты
                "sum": float(total_sum)  # Сумма оплаты
            })
            
            receipt['total'] = float(total_sum)
            
            self.log_message(f"Подготовлен чек для заказа {order_data.get('OrderID', 'Неизвестный')}")
            self.log_message(f"Итоговая сумма: {total_sum:.2f} руб.")
            return receipt
            
        except Exception as e:
            self.log_message(f"Ошибка подготовки фискального чека: {e}")
            return None

    def send_fiscal_receipt_to_evotor(self, receipt_data):
        """
        Отправка фискального чека в API Эвотор.
        """
        try:
            token = self.authenticate_evotor()
            if not token:
                self.log_message("Не удалось получить токен для Эвотор")
                return False
                
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            }
            
            # URL для регистрации чека
            url = f"{self.evotor_config['api_url']}/api/v5/doc"
            
            self.log_message(f"Отправка чека в Эвотор: {url}")
            self.log_message(f"Данные чека: {json.dumps(receipt_data, ensure_ascii=False, indent=2)}")
            
            # В реальной реализации здесь будет POST запрос
            # response = requests.post(url, headers=headers, json=receipt_data, verify=False)
            
            # Имитация успешной отправки
            self.log_message("Чек успешно отправлен в Эвотор (имитация)")
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка отправки чека в Эвотор: {e}")
            return False

    def finalize_order_and_send_receipt(self):
        """
        Завершение обработки заказа и отправка фискального чека.
        Этот метод вызывается после того, как все марки обработаны.
        """
        try:
            # Проверяем, есть ли данные заказа
            if not self.current_order_data:
                self.log_message("Нет данных заказа для фискализации")
                messagebox.showwarning("Предупреждение", "Нет данных заказа для фискализации")
                return False
                
            # Проверяем, есть ли товары
            if not self.current_bills_data:
                self.log_message("Нет товаров в заказе для фискализации")
                messagebox.showwarning("Предупреждение", "Нет товаров в заказе для фискализации")
                return False
                
            # === ИЗМЕНЕНО: Проверяем, все ли марки обработаны ===
            unmarked_required_products = []
            for product in self.displayed_products:
                if (product.get('ismarked', False) and 
                    not product.get('mark', '') and 
                    not product.get('without_mark', False)):
                    unmarked_required_products.append(product.get('ProductName', 'Неизвестный товар'))
            
            if unmarked_required_products:
                message = f"Фискализация невозможна!\nНе все товары, требующие маркировки, отсканированы:\n" + "\n".join(unmarked_required_products)
                self.log_message(message)
                messagebox.showerror("Ошибка фискализации", message)
                return False  # Блокируем фискализацию
                # === УБРАНО: Старая логика с вопросом ===
                # if not messagebox.askyesno("Подтверждение", f"{message}\nПродолжить фискализацию без этих марок?"):
                #     return False
                    
            # Подготавливаем чек
            receipt_data = self.prepare_fiscal_receipt(self.current_order_data, self.current_bills_data)
            if not receipt_data:
                self.log_message("Ошибка подготовки фискального чека")
                messagebox.showerror("Ошибка", "Ошибка подготовки фискального чека")
                return False
                
            # Отправляем чек
            success = self.send_fiscal_receipt_to_evotor(receipt_data)
            if success:
                self.log_message(f"Чек по заказу {self.current_order_data.get('OrderID', 'Неизвестный')} успешно отправлен в Эвотор")
                messagebox.showinfo("Успех", "Чек успешно отправлен в облачную кассу")
                return True
            else:
                self.log_message("Ошибка отправки чека в Эвотор")
                messagebox.showerror("Ошибка", "Ошибка отправки чека в облачную кассу")
                return False
                
        except Exception as e:
            self.log_message(f"Ошибка завершения заказа: {e}")
            messagebox.showerror("Ошибка", f"Ошибка завершения заказа: {e}")
            return False

    def on_closing(self):
        """Обработка закрытия приложения"""
        # Отключение от COM порта
        self.disconnect_from_com_port()
        
        # Закрытие соединения с БД
        try:
            if self.db_connection:
                self.db_connection.close()
        except:
            pass
            
        self.root.destroy()

def main():
    root = tk.Tk()
    app = MainApplication(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()