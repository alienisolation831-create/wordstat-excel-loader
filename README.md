# Wordstat Excel Loader

[English](#english) | [Русский](#russian)

<a id="english"></a>

## English

Wordstat Excel Loader is a self-hosted, single-user web application for collecting historical Yandex Wordstat search volumes from XLSX files. Synthetic DEMO mode is enabled by default. Real API requests require the owner's own credentials. Do not expose a running instance publicly without HTTPS, network protection and your own security audit.

### Quick start (Docker, local)

1. Copy `.env.example` to `.env`.
2. Set `APP_PASSWORD` to a unique random password (16–256 characters). Never commit `.env`.
3. Run `docker compose up --build -d` and open http://localhost:8000.
4. Log in, upload `example.xlsx`, click the file validation button (Проверить файл), select the worksheet, column and date range, then confirm the job. Download the Excel results from the jobs section.

DEMO data is synthetic and does not represent actual Yandex statistics. The application interface is currently in Russian.

To run locally without Docker, use Python 3.13, create a virtual environment with `python -m venv .venv`, activate it and run `pip install -r requirements.txt`. Set the environment variables from the example manually (do not pass the password as a command-line argument), then run:

```sh
uvicorn app:create_app --factory --host 127.0.0.1 --port 8000 --workers 1 --no-access-log
```

### Live API

Set `DEMO_MODE=false`, `YANDEX_API_KEY`, `YANDEX_FOLDER_ID`, `API_MAX_COST_KOPECKS` and `API_DAILY_BUDGET_KOPECKS` in the server's secret configuration. The service account needs the `search-api.webSearch.user` role, and the key needs the `yc.search-api.execute` scope.

The application uses one GetDynamics request per unique search phrase.

- [Official Wordstat API documentation](https://aistudio.yandex.ru/ru/docs/search-api/operations/wordstat-getdynamics)
- [Yandex Search API pricing](https://aistudio.yandex.ru/ru/docs/search-api/pricing)

**Data availability:** as documented on November 27, 2026, statistics after October 7, 2026 were unavailable. `WORDSTAT_MAX_DATE=2026-10-07` is a protective setting; the operator can change or clear it once data availability is confirmed. Do not request partial months for monthly aggregation. Weeks start on Monday and end on Sunday. Missing data points are not replaced with zeroes; totals include only available points.

### Server deployment (Amvera or VPS)

- `amvera.yaml` is provided as an example configuration. Compatibility with the current Amvera runtime has not been verified. Attach persistent storage at `/data` and set `DATA_DIR=/data`.
- Set `APP_ORIGIN=https://your-domain` and `COOKIE_SECURE=true`. Terminate TLS at a trusted reverse proxy that forwards the original Host header. Do not expose the backend port directly. Docker binds to loopback by default. Configure request-rate and body-size limits at the proxy and restrict access via VPN or IP allowlisting where possible.
- Run exactly one process and one instance with SQLite. Do not scale horizontally. When a job is stopped, the current request may finish, but new requests are not started. After a restart, an unfinished request is marked as failed rather than automatically retried and potentially charged again; remaining work continues.
- `MAX_REQUESTS_HOUR`, `MAX_REQUESTS_DAY` and `MAX_REQUESTS_RUN` limit request attempts. An estimated budget limit also applies. Configure budgets and quotas separately in Yandex Cloud, especially if the key is shared with other applications. API errors pause jobs; resuming does not automatically repeat the failed request.
- Data is stored locally in SQLite. Delete completed jobs through the interface and protect the storage volume and backups. Deletion does not guarantee cryptographic erasure. Do not share the application password.
- A new login invalidates the previous session. Five failed login attempts trigger a 15-minute block for the network address. Behind a reverse proxy without trusted client-address handling, all visitors may appear to have the same address. Restrict access and configure proxy rate limiting. Do not trust arbitrary X-Forwarded-For headers. All sessions are invalidated on restart, including a restart after a password change.

### Limits and budget

XLSX files must be no larger than 5 MB, with up to 5,000 rows, 50 columns and 400 characters per query. The date range is limited to 731 days, with up to 50 saved jobs and 100,000 potential daily data points per run.

Duplicates are removed. Formulas in the query column are unsupported. Workbooks containing macros or external links are rejected. The original workbook is not stored. Results are exported to two worksheets. Only text queries and single-user operation are supported. Short and long queries must currently be filtered in Excel.

Live mode requires `API_MAX_COST_KOPECKS` (a conservative upper estimate of the cost per request) and `API_DAILY_BUDGET_KOPECKS` (a rolling 24-hour budget), both expressed in kopecks. The request cap is the smaller of `MAX_REQUESTS_DAY` and the integer part of the budget divided by the estimated cost per request. This is an estimated safeguard, not a guarantee of actual charges. Account for pricing changes, taxes and other applications separately.

### Security testing

```sh
python -m unittest discover -s tests -v
```

These are local tests that do not require network access. Before public deployment, perform an up-to-date dependency and container audit, secret scanning, TLS and reverse proxy checks, OWASP ZAP testing and a limited live API test. Direct and known transitive dependencies are pinned for Python 3.13 on Linux. A complete audit and testing of the new Starlette version in a clean environment are required.

The project has not undergone independent penetration testing. Do not use it for confidential data without a separate audit and authorization to send queries to Yandex.

### Security update

Twenty tests passed locally with Starlette 0.50.0. Version 1.10.0 is specified to address known CVEs, but installing and testing that version was not possible in the development environment. Do not deploy a publicly accessible instance until CI checks and the dependency audit pass.

---

<a id="russian"></a>

## Русский

Самостоятельный однопользовательский веб-инструмент для сбора исторической частотности Яндекс Wordstat из XLSX. Синтетический DEMO включён по умолчанию, реальные запросы требуют отдельного ключа владельца. Не размещайте публичный рабочий экземпляр без HTTPS, сетевой защиты и собственного аудита.

## Быстрый старт (Docker, локально)
1. Скопируйте `.env.example` в `.env`.
2. Задайте `APP_PASSWORD` (16–256 символов, уникальный случайный пароль). Не добавляйте `.env` в Git.
3. `docker compose up --build -d`; откройте `http://localhost:8000`.
4. Войдите, загрузите `example.xlsx`, нажмите «Проверить файл», выберите лист, столбец, период, подтвердите запуск. Скачайте Excel в разделе задач. DEMO — вымышленные данные, не статистика Яндекса.

Локально без Docker: Python 3.13, `python -m venv .venv`, активируйте окружение, `pip install -r requirements.txt`, установите переменные окружения из примера вручную (не передавайте пароль аргументом командной строки), `uvicorn app:create_app --factory --host 127.0.0.1 --port 8000 --workers 1 --no-access-log`.

## Реальный API
Задайте `DEMO_MODE=false`, `YANDEX_API_KEY`, `YANDEX_FOLDER_ID`, `API_MAX_COST_KOPECKS`, `API_DAILY_BUDGET_KOPECKS` в секретах сервера. Сервисному аккаунту нужен `search-api.webSearch.user`, ключу — scope `yc.search-api.execute`. Приложение использует один запрос GetDynamics на уникальную фразу. API: https://aistudio.yandex.ru/ru/docs/search-api/operations/wordstat-getdynamics . Тарифы: https://aistudio.yandex.ru/ru/docs/search-api/pricing .

Внимание: на 27.11.2026 документация указывает отсутствие статистики после 07.10.2026. `WORDSTAT_MAX_DATE=2026-10-07` — защитная настройка; после подтверждения восстановления оператор может изменить или очистить её. Не запрашивайте неполные месяцы для месячной агрегации. Недельная начинается в понедельник и заканчивается в воскресенье. Пропущенные точки не подменяются нулями; сумма в итогах — только по доступным точкам.

## Сервер (Amvera или VPS)
- Для Amvera используется `amvera.yaml` как пример конфигурации; совместимость с текущим окружением платформы не проверена. Включите постоянный том `/data`, задайте `DATA_DIR=/data`.
- Обязательно `APP_ORIGIN=https://ваш-домен`, `COOKIE_SECURE=true`; TLS на доверенном reverse proxy, который передаёт оригинальный Host. Не открывайте порт backend напрямую. На VPS Docker по умолчанию привязан к loopback. Настройте лимиты запросов и тела на reverse proxy, ограничьте доступ VPN/IP, где возможно.
- Ровно один процесс и один экземпляр на SQLite. Не масштабируйте горизонтально. При остановке задачи текущий запрос может завершиться; новые не запускаются. После перезапуска незавершённый вызов отмечается ошибкой без повторной оплаты, остальные продолжаются.
- `MAX_REQUESTS_HOUR`, `MAX_REQUESTS_DAY`, `MAX_REQUESTS_RUN` ограничивают количество попыток; дополнительно действует оценочный бюджет. Настройте бюджет/квоты в Yandex Cloud отдельно, особенно если ключ используется другими приложениями. Ошибка API ставит задачу на паузу; продолжение не повторяет неудавшийся запрос автоматически.
- Данные хранятся локально в SQLite. Удаляйте завершённые задачи через интерфейс, защищайте том и резервные копии; удаление не гарантирует криптографического стирания. Используйте пароль, не передавайте его другим людям. Вход нового пользователя аннулирует старую сессию. После пяти неверных попыток блокируется сетевой адрес на 15 минут. За reverse proxy без доверенной настройки адреса все посетители могут иметь один адрес; ограничьте доступ VPN/IP и настройте rate limiting на proxy. Не доверяйте произвольным X-Forwarded-For. Все сессии аннулируются при перезапуске и смене пароля с перезапуском.

## Ограничения
XLSX ≤5 МБ, до 5000 строк и 50 столбцов, 400 символов в запросе, до 731 дня, 50 сохранённых задач, до 100 000 потенциальных дневных точек за запуск. Дубликаты удаляются. Формулы в столбце запросов не поддерживаются; XLSX с макросами/внешними ссылками отвергаются. Исходный Excel не хранится; результат выгружается в двух листах. Поддерживаются только текстовые запросы и один пользователь. Короткие/длинные запросы пока нужно фильтровать в Excel. В реальном режиме обязательны API_MAX_COST_KOPECKS (консервативная верхняя стоимость одного вызова) и API_DAILY_BUDGET_KOPECKS (бюджет на скользящие 24 часа). Лимит вызовов = минимум MAX_REQUESTS_DAY и целой части отношения бюджета к стоимости. Это оценочный предохранитель, не гарантия фактической суммы счёта: тарифы, налоги и другие приложения учитывайте отдельно.

## Проверка безопасности
`python -m unittest discover -s tests -v` (локальные тесты, без сети). Перед публичным запуском: актуальный аудит зависимостей и контейнера, проверка секретов, TLS, reverse proxy, OWASP ZAP и реальный ограниченный тест API. Версии прямых и известных транзитивных зависимостей зафиксированы для Python 3.13 Linux; полный аудит и тест новой версии Starlette в чистой среде обязательны. Проект не проходил независимый пентест. Не используйте его для конфиденциальных данных без отдельного аудита и разрешения на передачу запросов Яндексу.

## Security update
Локально проверено 20 тестов на Starlette 0.50.0; версия 1.10.0 указана для исправления известных CVE, но установка и тестирование этой версии в среде разработки были недоступны. До успешной проверки CI и аудита зависимостей не разворачивайте публичный экземпляр.
