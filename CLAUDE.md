# CLAUDE.md — уроки и правила работы над JewishGen_python

Читай этот файл в начале каждой сессии. Он содержит всё, что было выучено на практике.

---

## Проект

Python-скрапер для генеалогического поиска на FamilySearch (и MyHeritage).  
Стек: Playwright (async), PySide6 GUI, python-docx, openpyxl.

---

## Правила работы с git

- **Никогда не делай `git add -A` или `git add .`** — только конкретные файлы.
- **Перед коммитом** убедись что файл действительно содержит нужные изменения.
- **Revert — опасная операция**: `git revert HEAD` откатывает к последнему _закоммиченному_ состоянию, а не к тому что было в working copy до твоего коммита. Если до твоего коммита были незакоммиченные изменения пользователя — они войдут в твой коммит. Revert их сотрёт. В таком случае нужно делать `git revert HEAD` (revert of revert), чтобы вернуться к своему коммиту, а потом точечно откатить только свои правки.
- **Коммить только то, что просит пользователь.** Не коммить сам по себе.

---

## FamilySearch — алгоритм скрапера (строго соблюдать)

1. Открыть домашнюю страницу → заполнить First Names, Last Names (+ Place, Birth Year) → SEARCH.
2. Нажать таб Historical Records `[data-testid="hr-tab"]`. URL → `?tab=records`.
3. **Логин НЕ нужен** для показа строк таблицы HR tab. Строки видны без логина.
4. Собрать строки таблицы `tbody tr`. Если 0 строк — **не добавлять логин**, а увеличивать таймауты ожидания (XHR грузится медленно).
5. Открыть первую запись → если редирект на логин → залогиниться **один раз**.
6. Логин определяется по: `"login" in page.url` **ИЛИ** `"sign in" in page.title()` (JS-редирект может не завершиться до domcontentloaded).
7. После логина — `logged_in_ref[0] = True`. Повторный логин не делать.
8. Все записи открывать на **одной** странице (`page.goto(url)`), не в новых вкладках.
9. После каждой записи возвращаться на `results_url` через `page.goto(results_url)`.
10. Если есть Advanced Search — после первой записи: **обязательный reload**, HR tab, Advanced Search попап, заполнить поля, нажать Search, заново собрать результаты.

---

## FamilySearch — скачивание JPG

- **`expect_download` должен оборачивать ВЕСЬ блок**: первый клик на download кнопку + выбор JPG Only + клик confirm. Иначе если первый клик сам инициирует загрузку — событие пропускается.
- После клика на картинку — ждать `button[aria-label*="Download" i]` до 8 секунд перед тем как пробовать fallback-селекторы. После логина viewer грузится медленнее.
- Записи с `view=index` в URL — индексные, без изображения документа. Не вызывать `_download_jpg`, только сохранить thumbnail.
- Если `_best_img(page)` вернул пустую строку — не вызывать `_download_jpg` вообще.
- Таймаут `expect_download` — 45 секунд.
- Fallback если timeout: ждать 15 секунд появления файла в папке Downloads.

---

## FamilySearch — Advanced Search

### Открытие попапа
```
[data-testid="advanced-search-form-button"]
```
Нажать, ждать 1.5 секунды.

### Кнопка Search внутри попапа
```
[data-testid="search-button"]
```
Это `type="submit"`. Использовать именно этот testid, не `button:has-text("Search")`.

### Семейные члены — точные data-testid (из HTML)

| Родственник | Кнопка раскрытия | Given Name field | Given Exact | Surname field | Surname Exact |
|-------------|-----------------|-----------------|-------------|--------------|---------------|
| Spouse | `spouse-fieldGroupButton` | `spouseGivenName0-field` | `q_spouseGivenName_exact` | `spouseSurname0-field` | `q_spouseSurname_exact` |
| Father | `father-fieldGroupButton` | `fatherGivenName0-field` | `q_fatherGivenName_exact` | `fatherSurname0-field` | `q_fatherSurname_exact` |
| Mother | `mother-fieldGroupButton` | `motherGivenName0-field` | `q_motherGivenName_exact` | `motherSurname0-field` | `q_motherSurname_exact` |
| Other | `otherPerson-fieldGroupButton` | `otherGivenName0-field` | `q_otherGivenName_exact` | `otherSurname0-field` | `q_otherSurname_exact` |

**Важно**: плейсхолдеры у полей семейных членов **пустые** (`placeholder=""`). Искать только по `data-testid`, никогда по плейсхолдеру.

### Чекбоксы Exact
Кликать `[data-testid="q_spouseGivenName_exact"]` и т.д. напрямую. DOM-обход не нужен — testid известны.

### Поля событий (Birth/Marriage/Death/Residence)
```
input[placeholder*="City, County, State, Province, or Country"]
input[placeholder="Year"]
```
Перед заполнением кликнуть соответствующий таб (BIRTH / MARRIAGE / etc).

---

## FamilySearch — логин

- Форма: `#userName`, `#password`, `#login`
- Заполнять через `page.fill()` с retry до 3 раз (React-форма может сбрасывать значение)
- Ждать 2 секунды после появления `#userName` перед заполнением
- После `#login` — ждать редирект на `familysearch.org` без `login` в URL (таймаут 30 сек)

---

## Что НЕЛЬЗЯ делать

- **НЕ добавлять логин перед сбором результатов** — FamilySearch показывает строки таблицы без логина. Логин только при открытии первой записи.
- **НЕ использовать `.first` для полей семейных членов** — после раскрытия секции Spouse появляются новые поля с такими же плейсхолдерами как у предка. `.first` заполнит поля предка, а не супруга. Использовать `data-testid`.
- **НЕ использовать `button:has-text("Search")` для кнопки в попапе Advanced Search** — это может совпасть с другими кнопками. Использовать `[data-testid="search-button"]`.
- **НЕ делать поспешных выводов** о причине ошибки. "Строк: 0" — скорее всего медленный XHR, а не проблема с логином.

---

## GUI (PySide6)

- Advanced Search панель оборачивается в `QScrollArea` — если контент не влезает на экран, появляется скролл.
- Максимальная высота scroll area = 55% высоты экрана.
- Family Members секция: каждое поле (First Name, Last Name) имеет свой чекбокс "Exact".
- Автосохранение всех полей в `.fs_autosave.json` при каждом изменении.
- `_fam_fields[key]` хранит `(ff, cb_f, lf, cb_l)` — два инпута и два чекбокса.
