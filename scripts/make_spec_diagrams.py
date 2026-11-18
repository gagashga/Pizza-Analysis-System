import os
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, ArrowStyle

os.makedirs('docs/diagrams', exist_ok=True)
plt.rcParams['font.family'] = 'DejaVu Sans'

# 1. Conceptual Architecture Diagram
def create_conceptual_architecture():
    fig, ax = plt.subplots(figsize=(13, 8.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Title
    ax.text(50, 96, 'Концептуальная архитектура системы анализа «NewFastPizza»', 
            fontsize=15, fontweight='bold', ha='center', va='center', color='#1e293b')
    ax.text(50, 92.5, 'Высокоуровневое представление решения в терминах пользователей и бизнес-процессов', 
            fontsize=10, fontstyle='italic', ha='center', va='center', color='#64748b')

    # Layer 1: Users & Stakeholders (Top)
    users = [
        ('Менеджер сети', 'Анализ KPI, контроль очередей,\nисследование спроса K-Means', 15),
        ('Руководитель сети', 'Стратегические сводки,\nпринятие решений по филиалам', 38),
        ('Водитель-курьер', 'Получение оптимального филиала\nи кратчайшего маршрута', 62),
        ('ИТ-Администратор', 'Управление пользователями (RBAC),\nобновление дорожного графа', 85),
    ]

    for title, desc, cx in users:
        box = FancyBboxPatch((cx - 10, 77), 20, 12, boxstyle="round,pad=0.5,rounding_size=1.5",
                             facecolor='#f0fdf4', edgecolor='#16a34a', linewidth=1.5)
        ax.add_patch(box)
        ax.text(cx, 85.5, title, fontsize=9.5, fontweight='bold', ha='center', va='center', color='#166534')
        ax.text(cx, 80.5, desc, fontsize=7.5, ha='center', va='center', color='#374151')

    # Connectors from Users to Interface Layer
    for _, _, cx in users:
        ax.annotate('', xy=(cx, 69), xytext=(cx, 77),
                    arrowprops=dict(arrowstyle="-|>", color='#475569', lw=1.2))

    # Layer 2: Client Applications Layer
    client_box = FancyBboxPatch((5, 57), 90, 12, boxstyle="round,pad=0.5,rounding_size=2",
                                facecolor='#eff6ff', edgecolor='#2563eb', linewidth=1.8)
    ax.add_patch(client_box)
    ax.text(8, 66.5, 'УРОВЕНЬ КЛИЕНТСКИХ ИНТЕРФЕЙСОВ', fontsize=10, fontweight='bold', color='#1e40af')

    sub_clients = [
        ('Веб-дашборд менеджера и руководства\n(Vue 3, TypeScript, Pinia, Chart.js, Tailwind)', 25, 62, 34, 6.5),
        ('Интерактивные карты распределения\n(Зоны спроса K-Means, маршруты курьеров)', 58, 62, 28, 6.5),
        ('Клиент курьера / Мобильный шлюз\n(REST API / Мобильный терминал)', 83, 62, 20, 6.5)
    ]
    for title, cx, cy, w, h in sub_clients:
        b = FancyBboxPatch((cx - w/2, cy - h/2), w, h, boxstyle="round,pad=0.3,rounding_size=1",
                           facecolor='#ffffff', edgecolor='#3b82f6', linewidth=1)
        ax.add_patch(b)
        ax.text(cx, cy, title, fontsize=8, ha='center', va='center', color='#1e3a8a')

    # Arrow from Clients to Backend
    ax.annotate('', xy=(50, 49), xytext=(50, 57),
                arrowprops=dict(arrowstyle="-|>", color='#1e3a8a', lw=2, mutation_scale=15))
    ax.text(50.5, 53, 'Безопасный защищенный канал HTTPS / REST JSON API (JWT Bearer Token)', 
            fontsize=8.5, fontweight='bold', ha='center', va='center', color='#1e40af',
            bbox=dict(boxstyle='square,pad=0.2', facecolor='#ffffff', edgecolor='#93c5fd', lw=0.8))

    # Layer 3: Central Processing Platform (FastAPI Core)
    core_box = FancyBboxPatch((5, 20), 90, 29, boxstyle="round,pad=0.5,rounding_size=2",
                              facecolor='#f8fafc', edgecolor='#0f172a', linewidth=2)
    ax.add_patch(core_box)
    ax.text(8, 46.5, 'ЦЕНТРАЛЬНЫЙ АНАЛИТИЧЕСКИЙ СЕРВЕР «NewFastPizza» (FastAPI Modular Monolith)', 
            fontsize=10.5, fontweight='bold', color='#0f172a')

    modules = [
        ('Модуль авторизации и RBAC\n(JWT, OAuth2, Bcrypt,\nразграничение прав доступа)', 15, 34, 17, 18, '#fdf2f8', '#db2777', '#9d174d'),
        ('Аналитический модуль KPI\n(Расчет метрик заказов,\nнагрузка филиалов, динамика)', 33, 34, 16, 18, '#fff7ed', '#ea580c', '#9a3412'),
        ('Конструктор отчетов\n(Системные/пользовательские\nшаблоны, выгрузка в CSV)', 50, 34, 15, 18, '#fefce8', '#ca8a04', '#854d0e'),
        ('Интеллектуальная маршрутизация\n(Алгоритм Дейкстры на графе,\nучет пробок и очередей филиалов)', 68, 34, 18, 18, '#ecfdf5', '#059669', '#065f46'),
        ('Пространственный ML-анализ\n(Кластеризация K-Means,\nпоиск центроидов спроса)', 85, 34, 14, 18, '#f5f3ff', '#7c3aed', '#5b21b6'),
    ]

    for title, cx, cy, w, h, bg, border, tc in modules:
        b = FancyBboxPatch((cx - w/2, cy - h/2), w, h, boxstyle="round,pad=0.4,rounding_size=1",
                           facecolor=bg, edgecolor=border, linewidth=1.3)
        ax.add_patch(b)
        ax.text(cx, cy, title, fontsize=7.5, ha='center', va='center', color=tc, fontweight='bold')

    # Arrow from Backend to Data Layer
    ax.annotate('', xy=(30, 13), xytext=(30, 20),
                arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.8, mutation_scale=12))
    ax.annotate('', xy=(70, 13), xytext=(70, 20),
                arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.8, mutation_scale=12))

    # Layer 4: Data & Knowledge Layer
    db_box = FancyBboxPatch((8, 2), 42, 11, boxstyle="round,pad=0.4,rounding_size=1.5",
                            facecolor='#f1f5f9', edgecolor='#475569', linewidth=1.5)
    ax.add_patch(db_box)
    ax.text(29, 9.5, 'Реляционная СУБД (PostgreSQL 16)', fontsize=9.5, fontweight='bold', ha='center', color='#1e293b')
    ax.text(29, 5.5, 'Таблицы: users, branches, drivers, orders, report_templates\nТранзакции, внешние ключи, индексы, миграции Alembic', 
            fontsize=7.5, ha='center', color='#475569')

    graph_box = FancyBboxPatch((52, 2), 40, 11, boxstyle="round,pad=0.4,rounding_size=1.5",
                              facecolor='#f1f5f9', edgecolor='#475569', linewidth=1.5)
    ax.add_patch(graph_box)
    ax.text(72, 9.5, 'Цифровая модель дорог (NetworkX Graph)', fontsize=9.5, fontweight='bold', ha='center', color='#1e293b')
    ax.text(72, 5.5, 'Взвешенный ориентированный граф дорожной сети\n(узлы перекрестков, ребра дорог, веса времени с учетом пробок)', 
            fontsize=7.5, ha='center', color='#475569')

    plt.tight_layout()
    fig.savefig('docs/diagrams/conceptual_architecture.png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    print('Generated conceptual_architecture.png')

# 2. Activity Diagram: Driver Routing
def create_activity_driver_routing():
    fig, ax = plt.subplots(figsize=(10, 12), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 98, 'UML Диаграмма деятельности (Activity Diagram)', fontsize=13, fontweight='bold', ha='center', color='#0f172a')
    ax.text(50, 95.5, 'Вариант использования: «Расчет и выдача рекомендаций курьеру по выбору филиала»', fontsize=10.5, fontstyle='italic', ha='center', color='#334155')

    # Initial node
    start = plt.Circle((50, 91), 1.5, color='#0f172a')
    ax.add_patch(start)

    def draw_action(cx, cy, w, h, text, bg='#eff6ff', border='#2563eb', tc='#1e3a8a'):
        box = FancyBboxPatch((cx - w/2, cy - h/2), w, h, boxstyle="round,pad=0.3,rounding_size=1",
                             facecolor=bg, edgecolor=border, linewidth=1.4)
        ax.add_patch(box)
        ax.text(cx, cy, text, fontsize=8, ha='center', va='center', color=tc)
        return cy - h/2

    def draw_arrow(x1, y1, x2, y2, label=None, offset=(1, 0)):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.3, mutation_scale=10))
        if label:
            ax.text((x1+x2)/2 + offset[0], (y1+y2)/2 + offset[1], label, fontsize=7.5, color='#b91c1c', fontweight='bold')

    def draw_decision(cx, cy, size=4):
        diamond = patches.Polygon([[cx, cy + size], [cx + size*1.5, cy], [cx, cy - size], [cx - size*1.5, cy]],
                                  closed=True, facecolor='#fef9c3', edgecolor='#ca8a04', lw=1.4)
        ax.add_patch(diamond)
        return cy - size

    y = 91 - 1.5
    draw_arrow(50, y, 50, 84.5)
    
    # 1. Запрос
    b1_bottom = draw_action(50, 81.5, 60, 6, 'Курьер завершает доставку или запрашивает маршрут через приложение\n(передача текущих координат lat, lon курьера)')
    draw_arrow(50, b1_bottom, 50, 72.5)

    # 2. Валидация
    b2_bottom = draw_action(50, 69.5, 64, 6, 'API-шлюз валидирует координаты, проверяет JWT-токен курьера\nи обращается к GraphService для привязки к графу')
    draw_arrow(50, b2_bottom, 50, 62)

    # Decision 1
    d1_bottom = draw_decision(50, 59, size=3)
    ax.text(50, 59, 'Граф в RAM\nзагружен?', fontsize=6.5, ha='center', va='center', color='#713f12', fontweight='bold')
    
    # Yes -> down, No -> right branch
    draw_arrow(50, d1_bottom, 50, 48.5, label='[Да]')
    
    # No branch
    ax.plot([50 + 4.5, 82, 82, 50], [59, 59, 52, 52], color='#0f172a', lw=1.3)
    ax.annotate('', xy=(50, 48.5), xytext=(50, 52), arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.3))
    ax.text(62, 60, '[Нет: загрузка .pkl]', fontsize=7.5, color='#b91c1c', fontweight='bold')
    draw_action(82, 55.5, 26, 5, 'Загрузка графа из кэша\n`data/nn_graph.pkl`', bg='#fef2f2', border='#ef4444', tc='#991b1b')

    # 3. Поиск узла и расчет путей
    b3_bottom = draw_action(50, 45, 70, 7, 'GraphService находит ближайший узел графа к курьеру.\nЗапрос к BranchRepository: выборка всех активных филиалов и их очередей')
    draw_arrow(50, b3_bottom, 50, 36.5)

    # 4. Алгоритм Дейкстры
    b4_bottom = draw_action(50, 32.5, 72, 8, 'Алгоритм Дейкстры (RoutingService / NetworkX):\nРасчет кратчайших путей по времени с учетом весов ребер (текущие пробки).\nВычисление travel_time_minutes до каждого филиала')
    draw_arrow(50, b4_bottom, 50, 24)

    # 5. Многокритериальный выбор филиала
    b5_bottom = draw_action(50, 20, 72, 8, 'Многокритериальная оптимизация времени простоя курьера:\nScore = travel_time + W_queue * current_queue.\nВыбор целевого филиала с минимальным суммарным ожиданием')
    draw_arrow(50, b5_bottom, 50, 12)

    # 6. Формирование ответа
    b6_bottom = draw_action(50, 8.5, 72, 7, 'Формирование ответа DTO (id филиала, название, ETA, дистанция, геометрия пути).\nОтображение маршрута и рекомендаций на интерактивной карте курьера')
    draw_arrow(50, b6_bottom, 50, 2.5)

    # Final node (bullseye)
    end_outer = plt.Circle((50, 1.5), 1.3, color='#0f172a', fill=False, lw=1.5)
    end_inner = plt.Circle((50, 1.5), 0.8, color='#0f172a', fill=True)
    ax.add_patch(end_outer)
    ax.add_patch(end_inner)

    plt.tight_layout()
    fig.savefig('docs/diagrams/activity_driver_routing.png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    print('Generated activity_driver_routing.png')

# 3. Activity Diagram: K-Means Demand Analysis
def create_activity_demand_kmeans():
    fig, ax = plt.subplots(figsize=(10, 12), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 98, 'UML Диаграмма деятельности (Activity Diagram)', fontsize=13, fontweight='bold', ha='center', color='#0f172a')
    ax.text(50, 95.5, 'Вариант использования: «Анализ спроса и рекомендации по развитию сети филиалов (K-Means)»', fontsize=10, fontstyle='italic', ha='center', color='#334155')

    def draw_action(cx, cy, w, h, text, bg='#eff6ff', border='#2563eb', tc='#1e3a8a'):
        box = FancyBboxPatch((cx - w/2, cy - h/2), w, h, boxstyle="round,pad=0.3,rounding_size=1",
                             facecolor=bg, edgecolor=border, linewidth=1.4)
        ax.add_patch(box)
        ax.text(cx, cy, text, fontsize=8, ha='center', va='center', color=tc)
        return cy - h/2

    def draw_arrow(x1, y1, x2, y2, label=None, offset=(1, 0)):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.3, mutation_scale=10))
        if label:
            ax.text((x1+x2)/2 + offset[0], (y1+y2)/2 + offset[1], label, fontsize=7.5, color='#b91c1c', fontweight='bold')

    def draw_decision(cx, cy, size=3.5):
        diamond = patches.Polygon([[cx, cy + size], [cx + size*1.5, cy], [cx, cy - size], [cx - size*1.5, cy]],
                                  closed=True, facecolor='#fef9c3', edgecolor='#ca8a04', lw=1.4)
        ax.add_patch(diamond)
        return cy - size

    # Start
    start = plt.Circle((50, 91), 1.5, color='#0f172a')
    ax.add_patch(start)
    draw_arrow(50, 89.5, 50, 84.5)

    # 1. Параметры
    b1_bottom = draw_action(50, 81.5, 68, 6, 'Менеджер открывает экран «Зоны спроса K-Means» и задает параметры:\nпериод анализа дат, количество кластеров K, радиус обслуживания')
    draw_arrow(50, b1_bottom, 50, 73.5)

    # 2. Выборка
    b2_bottom = draw_action(50, 70.5, 68, 6, 'Backend (DemandAdapter) извлекает из БД массив выполненных заказов:\nгеографические координаты доставок (client_lat, client_lon)')
    draw_arrow(50, b2_bottom, 50, 63.5)

    # Decision: достаточно данных?
    d1_bottom = draw_decision(50, 60, size=3)
    ax.text(50, 60, 'Объем заказов\nN >= K?', fontsize=6.5, ha='center', va='center', color='#713f12', fontweight='bold')

    # [Нет] -> error exit
    ax.plot([50 + 4.5, 84, 84], [60, 60, 52], color='#0f172a', lw=1.3)
    draw_action(84, 49, 24, 6, 'Сообщение об ошибке:\n«Недостаточно данных»', bg='#fef2f2', border='#ef4444', tc='#991b1b')
    ax.plot([84, 84], [46, 40], color='#0f172a', lw=1.3)
    ax.annotate('', xy=(84, 38), xytext=(84, 40), arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.3))
    end_err_out = plt.Circle((84, 37), 1.2, color='#0f172a', fill=False, lw=1.4)
    end_err_in = plt.Circle((84, 37), 0.7, color='#0f172a', fill=True)
    ax.add_patch(end_err_out)
    ax.add_patch(end_err_in)
    ax.text(62, 61, '[Нет: N < K]', fontsize=7.5, color='#b91c1c', fontweight='bold')

    # [Да]
    draw_arrow(50, d1_bottom, 50, 50.5, label='[Да: N >= K]')

    # 3. K-Means
    b3_bottom = draw_action(50, 47, 72, 7, 'ML-пайплайн (Scikit-learn KMeans):\nКластеризация геопространственных точек заказов на K компактных зон.\nОпределение точных координат центроидов спроса и радиуса разброса')
    draw_arrow(50, b3_bottom, 50, 39)

    # 4. Сопоставление с филиалами
    b4_bottom = draw_action(50, 35, 72, 8, 'Пространственный анализ покрытия спроса:\nДля каждого центроида вычисляется расстояние до ближайшего филиала сети.\nОценка локальной плотности заказов относительно среднего уровня по городу')
    draw_arrow(50, b4_bottom, 50, 26.5)

    # 5. Классификация рекомендаций
    b5_bottom = draw_action(50, 22.5, 72, 8, 'Генерация стратегических рекомендаций по филиальной сети:\n• Если плотность спроса высокая, а филиала рядом нет -> РЕКОМЕНДАЦИЯ ОТКРЫТЬ ТОЧКУ\n• Если филиал имеет критически низкий спрос -> РЕКОМЕНДАЦИЯ ЗАКРЫТЬ/ОПТИМИЗИРОВАТЬ\n• Иначе -> Зона оптимально покрыта действующим филиалом')
    draw_arrow(50, b5_bottom, 50, 14.5)

    # 6. Рендеринг
    b6_bottom = draw_action(50, 10.5, 72, 8, 'Генерация интерактивной тепловой карты (Heatmap / GeoJSON) и карточек рекомендаций.\nОтображение результатов на дашборде менеджера с возможностью фильтрации')
    draw_arrow(50, b6_bottom, 50, 3)

    # Final
    end_outer = plt.Circle((50, 1.8), 1.2, color='#0f172a', fill=False, lw=1.5)
    end_inner = plt.Circle((50, 1.8), 0.7, color='#0f172a', fill=True)
    ax.add_patch(end_outer)
    ax.add_patch(end_inner)

    plt.tight_layout()
    fig.savefig('docs/diagrams/activity_demand_kmeans.png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    print('Generated activity_demand_kmeans.png')

# 4. Activity Diagram: Report Export
def create_activity_report_export():
    fig, ax = plt.subplots(figsize=(10, 11), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 98, 'UML Диаграмма деятельности (Activity Diagram)', fontsize=13, fontweight='bold', ha='center', color='#0f172a')
    ax.text(50, 95.5, 'Вариант использования: «Формирование и экспорт аналитического отчета»', fontsize=10.5, fontstyle='italic', ha='center', color='#334155')

    def draw_action(cx, cy, w, h, text, bg='#eff6ff', border='#2563eb', tc='#1e3a8a'):
        box = FancyBboxPatch((cx - w/2, cy - h/2), w, h, boxstyle="round,pad=0.3,rounding_size=1",
                             facecolor=bg, edgecolor=border, linewidth=1.4)
        ax.add_patch(box)
        ax.text(cx, cy, text, fontsize=8, ha='center', va='center', color=tc)
        return cy - h/2

    def draw_arrow(x1, y1, x2, y2, label=None, offset=(1, 0)):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.3, mutation_scale=10))
        if label:
            ax.text((x1+x2)/2 + offset[0], (y1+y2)/2 + offset[1], label, fontsize=7.5, color='#b91c1c', fontweight='bold')

    def draw_decision(cx, cy, size=3.5):
        diamond = patches.Polygon([[cx, cy + size], [cx + size*1.5, cy], [cx, cy - size], [cx - size*1.5, cy]],
                                  closed=True, facecolor='#fef9c3', edgecolor='#ca8a04', lw=1.4)
        ax.add_patch(diamond)
        return cy - size

    # Start
    start = plt.Circle((50, 91), 1.5, color='#0f172a')
    ax.add_patch(start)
    draw_arrow(50, 89.5, 50, 84.5)

    # 1
    b1_bottom = draw_action(50, 81.5, 68, 6, 'Менеджер выбирает шаблон отчета (системный или пользовательский)\nи настраивает фильтры: временной интервал, статус заказов, филиалы сети')
    draw_arrow(50, b1_bottom, 50, 73.5)

    # 2
    b2_bottom = draw_action(50, 70.5, 68, 6, 'Отправка запроса POST /api/v1/reports/custom с JSON-конфигурацией отчета.\nАвторизационный middleware проверяет права доступа пользователя')
    draw_arrow(50, b2_bottom, 50, 62.5)

    # 3
    b3_bottom = draw_action(50, 59, 72, 7, 'OrderRepository формирует параметризованный SQL-запрос к PostgreSQL:\nвыборка заказов по индексам ix_orders_order_time и ix_orders_branch_id_order_time')
    draw_arrow(50, b3_bottom, 50, 50.5)

    # 4
    b4_bottom = draw_action(50, 47, 72, 7, 'ReportService загружает данные в Pandas DataFrame и рассчитывает агрегаты:\nобщее число заказов, среднее время доставки, % отмен, распределение по часам/дням')
    draw_arrow(50, b4_bottom, 50, 38.5)

    # 5
    b5_bottom = draw_action(50, 35, 72, 7, 'Возврат сводных JSON-данных на клиентскую часть.\nРендеринг интерактивной таблицы и аналитических графиков (Chart.js / Recharts)')
    draw_arrow(50, b5_bottom, 50, 27)

    # Decision
    d1_bottom = draw_decision(50, 23.5, size=3)
    ax.text(50, 23.5, 'Менеджер нажал\n«Экспорт CSV»?', fontsize=6.5, ha='center', va='center', color='#713f12', fontweight='bold')

    # [Нет] -> direct to end
    ax.plot([50 - 4.5, 20, 20, 47], [23.5, 23.5, 3.5, 3.5], color='#0f172a', lw=1.3)
    ax.annotate('', xy=(47, 3.5), xytext=(43, 3.5), arrowprops=dict(arrowstyle="-|>", color='#0f172a', lw=1.3))
    ax.text(32, 24.5, '[Нет: только просмотр]', fontsize=7.5, color='#475569', fontweight='bold')

    # [Да]
    draw_arrow(50, d1_bottom, 50, 15, label='[Да: экспорт]')

    # 6
    b6_bottom = draw_action(50, 11, 68, 7, 'ReportService преобразует датафрейм в CSV-поток с BOM (UTF-8).\nКлиент получает файл с заголовком Content-Disposition: attachment; filename="report.csv"')
    draw_arrow(50, b6_bottom, 50, 4.5)

    # Final
    end_outer = plt.Circle((50, 3.5), 1.2, color='#0f172a', fill=False, lw=1.5)
    end_inner = plt.Circle((50, 3.5), 0.7, color='#0f172a', fill=True)
    ax.add_patch(end_outer)
    ax.add_patch(end_inner)

    plt.tight_layout()
    fig.savefig('docs/diagrams/activity_report_export.png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    print('Generated activity_report_export.png')

# 5. UML Class Diagram
def create_uml_class_diagram():
    fig, ax = plt.subplots(figsize=(15, 11), dpi=300)
    ax.set_xlim(0, 150)
    ax.set_ylim(0, 110)
    ax.axis('off')

    ax.text(75, 107.5, 'UML Диаграмма классов системы (Class Diagram)', fontsize=15, fontweight='bold', ha='center', color='#0f172a')
    ax.text(75, 104.5, 'Логический проект архитектуры: сущности БД, репозитории, сервисы бизнес-логики и контроллеры API', fontsize=9.5, fontstyle='italic', ha='center', color='#334155')

    def draw_class(x, y, w, h, title, stereotype=None, attrs=None, methods=None, bg='#ffffff', border='#1e3a8a'):
        box = FancyBboxPatch((x, y), w, h, boxstyle="square,pad=0", facecolor=bg, edgecolor=border, lw=1.2)
        ax.add_patch(box)
        
        # Header separator
        header_h = 4.5 if not stereotype else 6.0
        ax.plot([x, x+w], [y+h-header_h, y+h-header_h], color=border, lw=1)
        if stereotype:
            ax.text(x + w/2, y + h - 2.2, f'<<{stereotype}>>', fontsize=6.5, fontstyle='italic', ha='center', color='#475569')
            ax.text(x + w/2, y + h - 4.6, title, fontsize=8, fontweight='bold', ha='center', color='#0f172a')
        else:
            ax.text(x + w/2, y + h - 3, title, fontsize=8, fontweight='bold', ha='center', color='#0f172a')

        # Attributes
        cur_y = y + h - header_h - 2.2
        if attrs:
            for a in attrs:
                ax.text(x + 1.2, cur_y, a, fontsize=6.5, ha='left', color='#1e293b')
                cur_y -= 2.2
        
        # Method separator
        if methods:
            ax.plot([x, x+w], [cur_y + 1, cur_y + 1], color=border, lw=0.8)
            cur_y -= 1.2
            for m in methods:
                ax.text(x + 1.2, cur_y, m, fontsize=6.5, ha='left', color='#0369a1')
                cur_y -= 2.2

    # --- Package 1: Domain Entities (Bottom Left) ---
    p1 = FancyBboxPatch((4, 4), 68, 48, boxstyle="square,pad=0", facecolor='#f8fafc', edgecolor='#64748b', lw=1, linestyle='--')
    ax.add_patch(p1)
    ax.text(6, 49.5, 'Package: app.db.models (ORM Entities)', fontsize=8, fontweight='bold', color='#475569')

    draw_class(6, 26, 28, 22, 'User', stereotype='entity',
               attrs=['+ id: int [PK]', '+ username: str [UK]', '+ email: str [UK, null]', '+ password_hash: str', '+ role: str [enum]', '+ is_active: bool', '+ created_at: datetime'],
               methods=['+ verify_password()', '+ set_password()'])

    draw_class(40, 26, 30, 22, 'Driver', stereotype='entity',
               attrs=['+ id: int [PK]', '+ full_name: str', '+ phone: str [null]', '+ branch_id: int [FK]', '+ user_id: int [FK, UK]', '+ is_active: bool'],
               methods=['+ assign_branch()', '+ toggle_active()'])

    draw_class(6, 6, 28, 18, 'Branch', stereotype='entity',
               attrs=['+ id: int [PK]', '+ name: str [UK]', '+ lat: float', '+ lon: float', '+ current_queue: int', '+ is_active: bool'],
               methods=['+ update_queue()', '+ get_coordinates()'])

    draw_class(40, 6, 30, 18, 'Order', stereotype='entity',
               attrs=['+ id: int [PK]', '+ external_id: str [UK]', '+ client_lat, client_lon', '+ branch_id: int [FK]', '+ driver_id: int [FK]', '+ status: OrderStatus', '+ delivery_time_min: int'],
               methods=['+ mark_delivered()'])

    # --- Package 2: Repositories (Top Left) ---
    p2 = FancyBboxPatch((4, 55), 68, 46, boxstyle="square,pad=0", facecolor='#f0fdf4', edgecolor='#16a34a', lw=1, linestyle='--')
    ax.add_patch(p2)
    ax.text(6, 98.5, 'Package: app.repositories (Data Access)', fontsize=8, fontweight='bold', color='#166534')

    draw_class(6, 78, 30, 18, 'BaseRepository[T]', stereotype='repository',
               attrs=['# session: AsyncSession'],
               methods=['+ get_by_id(id): T', '+ list_all(): List[T]', '+ create(obj): T', '+ update(obj): T', '+ delete(id): bool'])

    draw_class(40, 78, 30, 18, 'OrderRepository', stereotype='repository',
               attrs=['- model: Order'],
               methods=['+ get_orders_in_range()', '+ get_branch_counts()', '+ get_delivery_kpis()', '+ create_batch()'])

    draw_class(6, 58, 30, 18, 'UserRepository', stereotype='repository',
               attrs=['- model: User'],
               methods=['+ get_by_username()', '+ get_by_email()', '+ list_paged()'])

    draw_class(40, 58, 30, 18, 'BranchRepository', stereotype='repository',
               attrs=['- model: Branch'],
               methods=['+ list_active()', '+ update_queue_val()', '+ get_with_orders()'])

    # --- Package 3: Business Services & ML (Center-Right) ---
    p3 = FancyBboxPatch((76, 4), 70, 56, boxstyle="square,pad=0", facecolor='#eff6ff', edgecolor='#2563eb', lw=1, linestyle='--')
    ax.add_patch(p3)
    ax.text(78, 57.5, 'Package: app.services & app.ml (Domain Services & ML Engine)', fontsize=8, fontweight='bold', color='#1e40af')

    draw_class(78, 38, 32, 17, 'StatsService', stereotype='service',
               attrs=['- order_repo: OrderRepo', '- branch_repo: BranchRepo'],
               methods=['+ calculate_kpis()', '+ get_branch_workload()', '+ get_daily_dynamics()'])

    draw_class(112, 38, 32, 17, 'ReportService', stereotype='service',
               attrs=['- stats_svc: StatsService', '- tmpl_repo: TemplateRepo'],
               methods=['+ generate_report_data()', '+ export_csv_stream()', '+ save_custom_template()'])

    draw_class(78, 20, 32, 16, 'RoutingAdapter', stereotype='service',
               attrs=['- routing_svc: RoutingService', '- branch_repo: BranchRepo'],
               methods=['+ find_best_branch()', '+ calculate_route_eta()'])

    draw_class(112, 20, 32, 16, 'DemandAdapter', stereotype='service',
               attrs=['- kmeans_analyzer: KMeans', '- order_repo: OrderRepo'],
               methods=['+ analyze_clusters()', '+ generate_branch_recs()'])

    draw_class(78, 6, 32, 12, 'AuthService', stereotype='service',
               attrs=['- user_repo: UserRepository'],
               methods=['+ authenticate_user()', '+ create_access_token()'])

    draw_class(112, 6, 32, 12, 'GraphService', stereotype='service',
               attrs=['- graph_state: GraphState'],
               methods=['+ load_graph_pickle()', '+ update_edge_weight()'])

    # --- Package 4: API Controllers (Top Right) ---
    p4 = FancyBboxPatch((76, 63), 70, 38, boxstyle="square,pad=0", facecolor='#fefce8', edgecolor='#ca8a04', lw=1, linestyle='--')
    ax.add_patch(p4)
    ax.text(78, 98.5, 'Package: app.api.v1.endpoints (REST Controllers)', fontsize=8, fontweight='bold', color='#854d0e')

    draw_class(78, 80, 32, 16, 'RecommendationsRouter', stereotype='controller',
               attrs=['/api/v1/recommendations'],
               methods=['+ get_driver_recommendation()', '+ get_demand_clusters()', '+ get_heatmap_layer()'])

    draw_class(112, 80, 32, 16, 'StatsRouter & ReportsRouter', stereotype='controller',
               attrs=['/api/v1/stats, /api/v1/reports'],
               methods=['+ get_summary_kpi()', '+ get_branch_stats()', '+ generate_custom_report()', '+ export_csv()'])

    draw_class(78, 65, 32, 13, 'AuthRouter', stereotype='controller',
               attrs=['/api/v1/auth'],
               methods=['+ login_access_token()', '+ register_user()', '+ read_users_me()'])

    draw_class(112, 65, 32, 13, 'AdminRouter', stereotype='controller',
               attrs=['/api/v1/admin'],
               methods=['+ list_system_users()', '+ update_role()', '+ reload_road_graph()'])

    # Connections & Associations
    # User -> Driver (1 to 0..1)
    ax.plot([34, 40], [37, 37], color='#1e3a8a', lw=1.2)
    ax.text(35, 38, '1', fontsize=6.5, color='#1e3a8a')
    ax.text(38.5, 38, '0..1', fontsize=6.5, color='#1e3a8a')

    # Branch -> Order (1 to *)
    ax.plot([34, 40], [15, 15], color='#1e3a8a', lw=1.2)
    ax.text(35, 16, '1', fontsize=6.5, color='#1e3a8a')
    ax.text(38.5, 16, '*', fontsize=6.5, color='#1e3a8a')

    # Driver -> Order (1 to *)
    ax.plot([55, 55], [26, 24], color='#1e3a8a', lw=1.2)
    ax.text(56, 25, '1..*', fontsize=6.5, color='#1e3a8a')

    # Inheritance OrderRepo -> BaseRepo
    ax.annotate('', xy=(21, 78), xytext=(40, 83),
                arrowprops=dict(arrowstyle="->", color='#16a34a', lw=1.2, linestyle=':'))
    
    # Controllers -> Services dependencies (dashed arrows)
    ax.annotate('', xy=(94, 55), xytext=(94, 63),
                arrowprops=dict(arrowstyle="->", color='#ca8a04', lw=1.2, linestyle='--'))
    ax.annotate('', xy=(128, 55), xytext=(128, 63),
                arrowprops=dict(arrowstyle="->", color='#ca8a04', lw=1.2, linestyle='--'))

    # Services -> Repositories dependencies (dashed arrows)
    ax.annotate('', xy=(70, 78), xytext=(78, 48),
                arrowprops=dict(arrowstyle="->", color='#2563eb', lw=1.2, linestyle='--'))

    plt.tight_layout()
    fig.savefig('docs/diagrams/uml_class_diagram.png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    print('Generated uml_class_diagram.png')

# 6. UML Deployment Diagram
def create_uml_deployment_diagram():
    fig, ax = plt.subplots(figsize=(14, 10.5), dpi=300)
    ax.set_xlim(0, 140)
    ax.set_ylim(0, 105)
    ax.axis('off')

    ax.text(70, 102, 'UML Диаграмма развертывания (Deployment Diagram)', fontsize=15, fontweight='bold', ha='center', color='#0f172a')
    ax.text(70, 99, 'Физический проект решения: распределение компонентов по узлам, сетевым протоколам и контейнерам', fontsize=9.5, fontstyle='italic', ha='center', color='#334155')

    def draw_node_3d(x, y, w, h, depth, title, stereotype='device', bg='#ffffff', border='#334155'):
        # Front face
        front = FancyBboxPatch((x, y), w, h, boxstyle="square,pad=0", facecolor=bg, edgecolor=border, lw=1.3)
        ax.add_patch(front)
        # Top face
        top = patches.Polygon([[x, y+h], [x+depth, y+h+depth], [x+w+depth, y+h+depth], [x+w, y+h]],
                              closed=True, facecolor='#e2e8f0', edgecolor=border, lw=1.1)
        ax.add_patch(top)
        # Right face
        right = patches.Polygon([[x+w, y], [x+w+depth, y+depth], [x+w+depth, y+h+depth], [x+w, y+h]],
                                closed=True, facecolor='#cbd5e1', edgecolor=border, lw=1.1)
        ax.add_patch(right)
        
        # Titles
        ax.text(x + w/2, y + h - 2.8, f'<<{stereotype}>>', fontsize=7, fontstyle='italic', ha='center', color='#475569')
        ax.text(x + w/2, y + h - 5.5, title, fontsize=8.5, fontweight='bold', ha='center', color='#0f172a')

    def draw_component(x, y, w, h, title, desc=None, bg='#eff6ff', border='#2563eb'):
        box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.2,rounding_size=0.8",
                             facecolor=bg, edgecolor=border, lw=1.2)
        ax.add_patch(box)
        # component tabs
        t1 = FancyBboxPatch((x - 1.2, y + h - 2.8), 2.4, 1.4, boxstyle="square,pad=0", facecolor='#ffffff', edgecolor=border, lw=1)
        t2 = FancyBboxPatch((x - 1.2, y + 1.4), 2.4, 1.4, boxstyle="square,pad=0", facecolor='#ffffff', edgecolor=border, lw=1)
        ax.add_patch(t1)
        ax.add_patch(t2)
        ax.text(x + w/2, y + h/2 + (1 if desc else 0), title, fontsize=7.5, fontweight='bold', ha='center', va='center', color='#1e3a8a')
        if desc:
            ax.text(x + w/2, y + h/2 - 2, desc, fontsize=6.5, ha='center', va='center', color='#475569')

    # --- Node 1: Client Workstation (Left) ---
    draw_node_3d(6, 48, 40, 42, 3, 'Рабочая станция менеджера / руководства\n(PC / Ноутбук под Windows / macOS / Linux)', bg='#f8fafc')
    draw_component(10, 68, 32, 11, 'Веб-браузер пользователя', 'Google Chrome / Firefox / Safari', bg='#ffffff', border='#3b82f6')
    draw_component(10, 52, 32, 12, 'SPA-клиент «NewFastPizza»', 'Vue 3, TypeScript, Pinia,\nChart.js, Leaflet / Tailwind', bg='#eff6ff', border='#2563eb')

    # --- Node 2: Driver Mobile Client (Bottom Left) ---
    draw_node_3d(6, 6, 40, 32, 3, 'Мобильный терминал курьера\n(Смартфон Android / iOS)', bg='#f8fafc')
    draw_component(10, 14, 32, 13, 'Клиентское приложение водителя', 'Мобильный браузер / REST-клиент\n(Передача GPS, прием маршрутов)', bg='#f0fdf4', border='#16a34a')

    # --- Node 3: Production VPS Host (Right) ---
    draw_node_3d(60, 4, 72, 88, 4, 'Целевой сервер: Production VPS Host\n(Linux Ubuntu 22.04 LTS, 4+ vCPU, 8+ GB RAM, 40 GB SSD)', bg='#f1f5f9')

    # Host level: Nginx Reverse Proxy
    draw_component(66, 75, 60, 11, 'Reverse Proxy / Web Server (Nginx)', 'Порты 80 HTTP (Redirect) / 443 HTTPS (TLS 1.3)\nРаздача статики SPA фронтенда + проксирование /api/*', bg='#fefce8', border='#ca8a04')

    # Inner box: Docker Environment
    docker_box = FancyBboxPatch((64, 8), 64, 62, boxstyle="square,pad=0", facecolor='#ffffff', edgecolor='#0284c7', lw=1.4, linestyle='--')
    ax.add_patch(docker_box)
    ax.text(66, 67.5, 'Docker Engine: Виртуальная сеть `pizza-network` (bridge 172.20.0.0/16)', fontsize=7.5, fontweight='bold', color='#0369a1')

    # Container 1: Backend API
    draw_node_3d(66, 27, 34, 37, 2, 'Контейнер: pizza-api\n(Docker: python:3.12-slim)', stereotype='executionEnvironment', bg='#f0fdf4', border='#16a34a')
    draw_component(68, 44, 30, 9, 'FastAPI REST Application', 'Uvicorn ASGI Server, порт 8000', bg='#ffffff', border='#16a34a')
    draw_component(68, 30, 30, 11, 'ML & Routing Core', 'Pandas, Scikit-learn (K-Means),\nNetworkX (Дейкстра), SQLAlchemy', bg='#ffffff', border='#16a34a')

    # Container 2: PostgreSQL
    draw_node_3d(104, 27, 22, 37, 2, 'Контейнер: pizza-pg\n(Postgres 16 Alpine)', stereotype='executionEnvironment', bg='#fdf2f8', border='#db2777')
    draw_component(106, 36, 18, 16, 'СУБД PostgreSQL 16', 'Внутренний порт 5432 TCP\nТаблицы, индексы, constraints', bg='#ffffff', border='#db2777')

    # Persistent Volumes
    v1 = FancyBboxPatch((68, 11), 28, 9, boxstyle="round,pad=0.2,rounding_size=0.6", facecolor='#f8fafc', edgecolor='#64748b', lw=1)
    ax.add_patch(v1)
    ax.text(82, 16.5, '<<storage volume>> app_data', fontsize=6.5, fontweight='bold', ha='center', color='#334155')
    ax.text(82, 13, 'Кэш графа дорог (nn_graph.pkl), логи', fontsize=6, ha='center', color='#64748b')

    v2 = FancyBboxPatch((104, 11), 22, 9, boxstyle="round,pad=0.2,rounding_size=0.6", facecolor='#f8fafc', edgecolor='#64748b', lw=1)
    ax.add_patch(v2)
    ax.text(115, 16.5, '<<volume>> postgres_data', fontsize=6.5, fontweight='bold', ha='center', color='#334155')
    ax.text(115, 13, '/var/lib/postgresql/data', fontsize=6, ha='center', color='#64748b')

    # Network connections
    # Client -> Nginx
    ax.plot([46, 66], [74, 80], color='#2563eb', lw=1.5)
    ax.text(53, 79, '<<HTTPS>> :443\nREST JSON', fontsize=6.5, color='#1e40af', fontweight='bold', ha='center')

    # Driver -> Nginx
    ax.plot([46, 66], [22, 75], color='#16a34a', lw=1.5)
    ax.text(53, 44, '<<HTTPS>> :443\nREST / Recommendations', fontsize=6.5, color='#166534', fontweight='bold', ha='center')

    # Nginx -> pizza-api
    ax.annotate('', xy=(80, 64), xytext=(80, 75),
                arrowprops=dict(arrowstyle="<|-|>", color='#ca8a04', lw=1.4))
    ax.text(81, 70, '<<HTTP reverse proxy>> :8000', fontsize=6.5, color='#854d0e', fontweight='bold')

    # pizza-api -> pizza-pg
    ax.annotate('', xy=(104, 45), xytext=(100, 45),
                arrowprops=dict(arrowstyle="<|-|>", color='#0f172a', lw=1.4))
    ax.text(102, 47, '<<TCP>> :5432\nSQL Pool', fontsize=6, color='#0f172a', fontweight='bold', ha='center')

    # Volume mounts
    ax.annotate('', xy=(82, 27), xytext=(82, 20),
                arrowprops=dict(arrowstyle="-", color='#64748b', lw=1.2, linestyle=':'))
    ax.annotate('', xy=(115, 27), xytext=(115, 20),
                arrowprops=dict(arrowstyle="-", color='#64748b', lw=1.2, linestyle=':'))

    plt.tight_layout()
    fig.savefig('docs/diagrams/uml_deployment_diagram.png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    print('Generated uml_deployment_diagram.png')

if __name__ == '__main__':
    create_conceptual_architecture()
    create_activity_driver_routing()
    create_activity_demand_kmeans()
    create_activity_report_export()
    create_uml_class_diagram()
    create_uml_deployment_diagram()
    print('All diagrams generated successfully!')

