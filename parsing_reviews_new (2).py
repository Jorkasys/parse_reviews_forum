import time
from typing import List, Dict, Set, Optional, Tuple
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import re
from datetime import datetime, timedelta
import hashlib
import os
import tempfile
from dotenv import load_dotenv
import boto3
from botocore.config import Config
import pandas as pd
from openpyxl import Workbook
import uuid
import shutil


load_dotenv()

# Параметры из окружения
Access_key = os.getenv("Access_key")
Secret_key = os.getenv("Secret_key")
Region = os.getenv("Region")
Endpoint = os.getenv("Endpoint")
Department_busket_name = os.getenv("Department_busket_name")

required_vars = {
    "Access_key": Access_key,
    "Secret_key": Secret_key,
    "Region": Region,
    "Endpoint": Endpoint,
    "Department_busket_name": Department_busket_name,
}
for name, value in required_vars.items():
    if not value:
        raise EnvironmentError(f"Переменная окружения {name} не задана")


def upload_to_cloud_sync(file_path: str) -> bool:
    """Синхронная загрузка файла в облако"""
    try:
        s3_client = boto3.client(
            's3',
            aws_access_key_id=Access_key,
            aws_secret_access_key=Secret_key,
            region_name=Region,
            endpoint_url=Endpoint,
            config=Config(
                signature_version='s3v4'
            )
        )

        file_name = os.path.basename(file_path)

        s3_client.upload_file(
            file_path,
            Department_busket_name,
            file_name,
            ExtraArgs={'ContentType': 'application/octet-stream'}
        )

        print(f"  ✅ Загружено в облако: {file_name}")
        return True

    except Exception as e:
        print(f"  ❌ Ошибка загрузки в облако: {e}")
        return False


def delete_local_file(file_path: str):
    """Удаление локального файла после загрузки"""
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"  🗑️ Удален локальный файл: {os.path.basename(file_path)}")
    except Exception as e:
        print(f"  ⚠️ Не удалось удалить файл {file_path}: {e}")


def download_latest_from_cloud(prefix: str = 'reviews_transformed_') -> Optional[str]:
    """Скачивание последнего файла из облака"""
    try:
        s3_client = boto3.client(
            's3',
            aws_access_key_id=Access_key,
            aws_secret_access_key=Secret_key,
            region_name=Region,
            endpoint_url=Endpoint,
            config=Config(
                signature_version='s3v4'
            )
        )

        response = s3_client.list_objects_v2(
            Bucket=Department_busket_name,
            Prefix=prefix
        )

        if 'Contents' not in response:
            print(f"📦 В облаке нет файлов с префиксом '{prefix}'")
            return None

        parquet_files = [
            obj['Key'] for obj in response['Contents']
            if obj['Key'].endswith('.parquet') and prefix in obj['Key']
        ]

        if not parquet_files:
            print("📦 В облаке нет parquet файлов")
            return None

        latest_file = sorted(parquet_files, reverse=True)[0]

        print(f"📥 Скачивание последнего файла из облака: {latest_file}")

        local_filename = "temp_downloaded.parquet"
        s3_client.download_file(Department_busket_name, latest_file, local_filename)

        print(f"✅ Загружено из облака: {latest_file}")
        print(f"   → Локально сохранено как: {local_filename}")

        return local_filename

    except Exception as e:
        print(f"❌ Ошибка скачивания из облака: {e}")
        return None


def create_excel_with_retry(df, filename, max_retries=3):
    """Создает Excel файл с повторными попытками"""
    for attempt in range(max_retries):
        try:
            temp_filename = f"temp_{filename}"

            # ИСПРАВЛЕНИЕ: заменяем pd.NA на None для совместимости с Excel
            df_excel = df.copy()

            # Заменяем pd.NA на None во всех колонках
            for col in df_excel.columns:
                if df_excel[col].dtype == 'Int64':
                    df_excel[col] = df_excel[col].astype('object')  # Конвертируем в object
                    df_excel[col] = df_excel[col].where(df_excel[col].notna(), None)  # NA -> None

            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"

            # Записываем заголовки
            headers = list(df_excel.columns)
            ws.append(headers)

            # Записываем данные
            for _, row in df_excel.iterrows():
                # Преобразуем datetime в строку для Excel
                row_values = []
                for val in row:
                    if pd.isna(val):
                        row_values.append(None)
                    elif isinstance(val, datetime):
                        row_values.append(val.strftime('%Y-%m-%d %H:%M:%S'))
                    else:
                        row_values.append(val)
                ws.append(row_values)

            # Настраиваем ширину колонок
            widths = {
                'A': 35,  # id (MD5 hash)
                'B': 20,  # created_date
                'C': 25,  # hub
                'D': 20,  # category_name
                'E': 50,  # url
                'F': 15,  # type
                'G': 40,  # title
                'H': 80,  # content
                'I': 12,  # views_count
                'J': 12,  # like_count
                'K': 12,  # repost_count
                'L': 15,  # comments_count
                'M': 10,  # rating
            }

            for col, width in widths.items():
                ws.column_dimensions[col].width = width

            # Сохраняем во временный файл
            wb.save(temp_filename)

            # Переименовываем в финальный
            if os.path.exists(filename):
                os.remove(filename)
            os.rename(temp_filename, filename)

            print(f"✅ Excel создан: {filename}")
            return True

        except Exception as e:
            print(f"⚠️ Попытка {attempt + 1}/{max_retries} не удалась: {e}")
            time.sleep(1)

    return False


class YandexFinanceSeleniumParser:
    def __init__(self, headless=False, use_cloud=False):
        """
        При инициализации скачивается последний файл из облака (если use_cloud=True)

        Args:
            headless: запуск браузера в headless режиме
            use_cloud: использовать облачное хранилище (False для локального тестирования)
        """
        self.driver = None
        self.headless = headless
        self.use_cloud = use_cloud

        # Генерируем имя файла с ТЕКУЩЕЙ датой (для сохранения)
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.database_file = f"reviews_transformed_{current_date}.parquet"

        self.existing_reviews = {}
        self.last_review_dates = {}
        self.existing_ids = set()

        print(f"📦 Целевой файл: {self.database_file}")
        print(f"☁️ Режим облака: {'ВКЛЮЧЕН' if use_cloud else 'ОТКЛЮЧЕН (локальное тестирование)'}")

        # Пытаемся скачать последний файл из облака
        if self.use_cloud:
            print("🔍 Поиск последней версии базы данных в облаке...")
            downloaded_file = download_latest_from_cloud()

            if downloaded_file:
                try:
                    shutil.move(downloaded_file, self.database_file)  # ← ИСПРАВЛЕНО
                    print(f"✅ База данных готова к работе: {self.database_file}")
                except Exception as e:
                    print(f"⚠️ Ошибка переименования: {e}")
                    shutil.copy2(downloaded_file, self.database_file)
                    os.remove(downloaded_file)
                    print(f"✅ База данных скопирована: {self.database_file}")
            else:
                print("📝 Будет создана новая база данных")

        try:
            self.setup_driver()
        except Exception as e:
            print(f"❌ Критическая ошибка инициализации Chrome: {e}")
            print("💡 Проверьте:")
            print("   1. Установлен ли Google Chrome")
            print("   2. Установлен ли chromedriver")
            print("   3. Совпадают ли версии Chrome и chromedriver")
            raise RuntimeError("Не удалось инициализировать Selenium WebDriver") from e


    def generate_review_id(self, content: str, url: str) -> str:
        """Генерация ID как MD5(content + url) согласно требованиям"""
        unique_string = f"{content}{url}"
        return hashlib.md5(unique_string.encode()).hexdigest()

    def setup_driver(self):
        """Настройка драйвера Chrome"""
        # Создаем УНИКАЛЬНУЮ временную папку
        temp_profile_dir = os.path.join(
            tempfile.gettempdir(),
            f"chrome_profile_{uuid.uuid4().hex[:8]}"
        )

        # Удаляем если существует (остатки от прошлого запуска)
        if os.path.exists(temp_profile_dir):
            try:
                shutil.rmtree(temp_profile_dir)
            except:
                pass

        # Создаем заново
        os.makedirs(temp_profile_dir, exist_ok=True)

        options = Options()

        # Настройки для обхода обнаружения автоматизации
        options.add_argument(f"--user-data-dir={temp_profile_dir}")
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # Основные настройки
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')

        # Для Linux - дополнительные флаги
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--disable-extensions')

        options.add_argument(
            'user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )

        if self.headless:
            options.add_argument('--headless=new')

        try:
            self.driver = webdriver.Chrome(options=options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            self.temp_profile_dir = temp_profile_dir

        except Exception as e:
            print(f"❌ Ошибка инициализации Chrome: {e}")
            # Очищаем временную папку при ошибке
            if os.path.exists(temp_profile_dir):
                try:
                    shutil.rmtree(temp_profile_dir)
                except:
                    pass
            raise

    def close(self):
        """Закрытие драйвера"""
        if self.driver:
            self.driver.quit()

        # Очищаем временную папку профиля
        if hasattr(self, 'temp_profile_dir') and os.path.exists(self.temp_profile_dir):
            try:
                shutil.rmtree(self.temp_profile_dir)
                print(f"🗑️ Очищена временная папка Chrome")
            except Exception as e:
                print(f"⚠️ Не удалось очистить временную папку: {e}")

    def generate_review_hash(self, review_data: Dict) -> str:
        """Генерация уникального хеша для отзыва"""
        unique_string = f"{review_data.get('product', '')}{review_data.get('author', '')}{review_data.get('text', '')}{review_data.get('date', '')}"
        return hashlib.md5(unique_string.encode()).hexdigest()

    def load_database(self) -> Tuple[Set[str], Dict[str, datetime]]:
        """Загрузка существующих ID из базы данных"""
        existing_ids = set()
        last_dates = {}

        if not os.path.exists(self.database_file):
            print(f"📦 База данных {self.database_file} будет создана при первом сохранении")
            return existing_ids, last_dates

        try:
            df = pd.read_parquet(self.database_file)

            print(f"📂 Загружена база данных: {self.database_file}")
            print(f"   Всего записей: {len(df)}")

            if 'id' in df.columns:
                existing_ids = set(df['id'].unique())
                print(f"   Уникальных ID: {len(existing_ids)}")

            if 'url' in df.columns and 'created_date' in df.columns:
                for url in df['url'].unique():
                    if pd.notna(url) and 'reviews/' in url:
                        product = url.split('reviews/')[-1]
                        product_df = df[df['url'] == url]
                        dates = pd.to_datetime(product_df['created_date'], errors='coerce')
                        latest_date = dates.max()

                        if pd.notna(latest_date):
                            last_dates[product] = latest_date
                            print(f"   {product}: {len(product_df)} отзывов, последний от {latest_date:%Y-%m-%d}")

        except Exception as e:
            print(f"❌ Ошибка при загрузке базы данных: {e}")

        return existing_ids, last_dates

    def is_review_old(self, review_date_str: str, product: str) -> bool:
        """Проверка, является ли отзыв старым"""
        if product not in self.last_review_dates:
            return False

        try:
            review_datetime = self.parse_date(review_date_str)
            last_known_date = self.last_review_dates[product]
            return review_datetime <= last_known_date
        except:
            return False

    def is_review_duplicate(self, review_data: Dict, product: str) -> bool:
        """Проверка на дубликат по хешу"""
        review_hash = self.generate_review_hash(review_data)
        product_hashes = self.existing_reviews.get(product, set())
        return review_hash in product_hashes

    def wait_for_reviews(self, timeout=10):
        """Ожидание загрузки отзывов"""
        try:
            wait = WebDriverWait(self.driver, timeout)

            selectors_to_try = [
                (By.CLASS_NAME, "Review-Text"),
                (By.CLASS_NAME, "Cut-Visible"),
                (By.CSS_SELECTOR, "[class*='Review']"),
            ]

            for by, selector in selectors_to_try:
                try:
                    elements = wait.until(EC.presence_of_all_elements_located((by, selector)))
                    if elements:
                        return True
                except TimeoutException:
                    continue

            return False
        except Exception as e:
            print(f"⚠️ Ошибка при ожидании отзывов: {e}")
            return False

    def check_for_old_reviews_quick(self, product: str) -> bool:
        """Быстрая проверка первых отзывов на странице"""
        if product not in self.last_review_dates:
            return False

        try:
            date_elements = self.driver.find_elements(By.CLASS_NAME, "Review-Date")[:5]

            for date_elem in date_elements:
                date_text = date_elem.text
                if self.is_review_old(date_text, product):
                    print(f"  ⚠️ Обнаружен старый отзыв от {date_text} - остановка")
                    return True

            return False
        except:
            return False

    def smart_scroll(self, product: str, max_scrolls: int = 100) -> int:
        """
        Умная прокрутка с РАННЕЙ ОСТАНОВКОЙ при обнаружении старых отзывов
        """
        is_update_mode = product in self.last_review_dates

        if is_update_mode:
            print(
                f"  🔄 Режим обновления для {product} (последний известный: {self.last_review_dates[product]:%Y-%m-%d})")
            actual_max_scrolls = min(5, max_scrolls)
        else:
            print(f"  🆕 Первый парсинг {product} - полная загрузка")
            actual_max_scrolls = max_scrolls

        scrolls_done = 0
        consecutive_old_reviews = 0
        last_height = self.driver.execute_script("return document.body.scrollHeight")

        for scroll in range(actual_max_scrolls):
            if is_update_mode:
                date_elements = self.driver.find_elements(By.CLASS_NAME, "Review-Date")[-10:]
                old_count = 0

                for date_elem in date_elements:
                    try:
                        date_text = date_elem.text
                        if self.is_review_old(date_text, product):
                            old_count += 1
                    except:
                        continue

                if old_count >= 5:
                    consecutive_old_reviews += 1
                    if consecutive_old_reviews >= 2:
                        print(f"  🛑 ОСТАНОВКА: достигнуты старые отзывы (прокрутка {scroll})")
                        break
                else:
                    consecutive_old_reviews = 0

            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1.5)

            try:
                show_more_btn = self.driver.find_element(By.CSS_SELECTOR,
                                                         "button[class*='More'], button[class*='show']")
                if show_more_btn.is_displayed() and show_more_btn.is_enabled():
                    show_more_btn.click()
                    print(f"  ✓ Показать еще (прокрутка {scroll + 1})")
                    time.sleep(1.5)
            except (NoSuchElementException, StaleElementReferenceException):
                pass
            except Exception:
                pass

            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                print(f"  📄 Конец страницы (прокрутка {scroll + 1})")
                break
            last_height = new_height
            scrolls_done += 1

        print(f"  ✅ Прокруток выполнено: {scrolls_done}")
        return scrolls_done

    def extract_reviews_with_required_format(self, product: str) -> List[Dict]:
        """Извлечение отзывов в ТРЕБУЕМОМ ФОРМАТЕ с РАННЕЙ ОСТАНОВКОЙ"""
        reviews = []
        is_update_mode = product in self.last_review_dates

        self.expand_reviews_selectively(product)
        base_url = self.driver.current_url

        soup = BeautifulSoup(self.driver.page_source, 'html.parser')
        review_texts = soup.find_all('div', class_='Review-Text')

        print(f"📝 Обработка {len(review_texts)} отзывов...")

        consecutive_duplicates = 0
        consecutive_old = 0

        for idx, review_text_elem in enumerate(review_texts, 1):
            try:
                parent = review_text_elem.parent
                while parent and not any(cls in parent.get('class', []) for cls in ['Review', 'ReviewItem']):
                    parent = parent.parent
                    if parent and parent.name == 'body':
                        parent = review_text_elem.parent
                        break

                if not parent:
                    parent = review_text_elem.parent

                if is_update_mode:
                    date_elem = parent.find('a', class_='Review-Date')
                    if not date_elem:
                        date_elem = parent.find('span', class_='Review-Date')

                    if date_elem:
                        date_text = date_elem.get_text(strip=True)
                        if self.is_review_old(date_text, product):
                            consecutive_old += 1
                            if consecutive_old >= 10:
                                print(f"  🛑 ОСТАНОВКА: 10 старых отзывов подряд")
                                break
                            continue
                        else:
                            consecutive_old = 0

                text_parts = []

                text_cut = review_text_elem.find('div', class_='TextCut')
                if text_cut:
                    text_parts.append(text_cut.get_text(strip=True))

                visible_text = review_text_elem.find('span', class_='Cut-Visible')
                if visible_text:
                    text_parts.append(visible_text.get_text(strip=True))

                invisible_text = review_text_elem.find('span', class_='Cut-Invisible')
                if invisible_text:
                    for hide_elem in invisible_text.find_all('span', class_='Cut-Hide'):
                        hide_elem.decompose()
                    invisible_content = invisible_text.get_text(strip=True)
                    if invisible_content:
                        text_parts.append(invisible_content)

                content = " ".join(text_parts) if text_parts else review_text_elem.get_text(strip=True)

                if not content or len(content) < 10:
                    continue
                if content.strip().isdigit():
                    continue
                if content in ['Показать ещё', 'Скрыть']:
                    continue

                review_id = self.generate_review_id(content, base_url)

                if review_id in self.existing_ids:
                    consecutive_duplicates += 1
                    if consecutive_duplicates >= 10:
                        print(f"  🛑 ОСТАНОВКА: 10 дубликатов подряд")
                        break
                    continue
                else:
                    consecutive_duplicates = 0

                created_date = datetime.now()
                date_elem = parent.find('a', class_='Review-Date')
                if not date_elem:
                    date_elem = parent.find('span', class_='Review-Date')
                if date_elem:
                    date_text = date_elem.get_text(strip=True)
                    created_date = self.parse_date(date_text)

                comments_count = None
                author_elem = parent.find('span', class_='ReviewAuthor-ExtraInfoReviewsCount')
                if author_elem:
                    reviews_count_match = re.search(r'(\d+)', author_elem.get_text())
                    if reviews_count_match:
                        comments_count = int(reviews_count_match.group(1))

                rating = None
                rating_container = parent.find('div', class_='Review-Rating')
                if rating_container:
                    stars = rating_container.find_all('span', class_='Review-RatingStar')
                    filled_stars = 0
                    for star in stars:
                        if 'Review-RatingStar_view_full' in ' '.join(star.get('class', [])):
                            filled_stars += 1
                    if filled_stars > 0:
                        rating = float(filled_stars)

                like_count = None
                reactions_elem = parent.find('div', class_='Reactions')
                if reactions_elem:
                    like_button = reactions_elem.find('button', attrs={'aria-label': re.compile(r'\d+ лайк')})
                    if like_button:
                        like_match = re.search(r'(\d+)', like_button.get('aria-label', ''))
                        if like_match:
                            like_count = int(like_match.group(1))

                title = f"Отзыв о {product.replace('_', ' ').title()}"

                review_data = {
                    'id': review_id,
                    'created_date': created_date,
                    'hub': 'yandex.ru/finance',
                    'category_name': 'Агрегатор отзывов',
                    'url': base_url,
                    'type': 'отзыв',
                    'title': title,
                    'content': content,
                    'views_count': None,
                    'like_count': like_count,
                    'repost_count': None,
                    'comments_count': comments_count,
                    'rating': rating
                }

                reviews.append(review_data)
                self.existing_ids.add(review_id)

                if len(reviews) <= 3:
                    print(f"  ✅ Новый: отзыв #{idx}")
                elif idx % 50 == 0:
                    print(f"  Обработано: {idx}/{len(review_texts)}")

            except Exception as e:
                continue

        if is_update_mode and (consecutive_duplicates > 0 or consecutive_old > 0):
            print(f"  ⏭️ Пропущено: {consecutive_duplicates} дубликатов, {consecutive_old} старых")

        print(f"✅ Извлечено {len(reviews)} {'НОВЫХ' if is_update_mode else ''} отзывов")
        return reviews

    def expand_reviews_selectively(self, product: str):
        """СЕЛЕКТИВНОЕ разворачивание - только новых отзывов"""
        is_update_mode = product in self.last_review_dates

        if not is_update_mode:
            print("  📖 Разворачиваем ВСЕ тексты (первый парсинг)...")
            max_expand = 1000
        else:
            print("  📖 Разворачиваем только НОВЫЕ тексты...")
            max_expand = 20

        expanded = 0
        buttons = self.driver.find_elements(By.CSS_SELECTOR, 'span.Link_theme_ghost[role="button"]')

        for btn in buttons[:max_expand]:
            try:
                if not btn.is_displayed():
                    continue

                if "Читать ещё" not in btn.text:
                    continue

                if is_update_mode:
                    try:
                        parent = btn.find_element(By.XPATH, "./ancestor::div[contains(@class, 'Review')]")
                        date_elem = parent.find_element(By.CLASS_NAME, "Review-Date")
                        date_text = date_elem.text
                        if self.is_review_old(date_text, product):
                            continue
                    except (NoSuchElementException, StaleElementReferenceException):
                        pass  # Не удалось найти дату, все равно разворачиваем

                self.driver.execute_script("arguments[0].click();", btn)
                expanded += 1

                if expanded % 10 == 0:
                    time.sleep(0.3)

            except StaleElementReferenceException:
                continue  # Элемент устарел, пропускаем
            except Exception:
                continue  # Любая другая ошибка

        if expanded > 0:
            print(f"  ✓ Развернуто {expanded} отзывов")

    def parse_date(self, date_text: str) -> datetime:
        """Преобразование даты в datetime объект"""
        months = {
            'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4,
            'мая': 5, 'июня': 6, 'июля': 7, 'августа': 8,
            'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12
        }

        current_year = datetime.now().year

        for month_name, month_num in months.items():
            if month_name in date_text:
                day_match = re.search(r'(\d+)', date_text)
                if day_match:
                    day = int(day_match.group(1))
                    return datetime(current_year, month_num, day)

        return datetime.now()

    def update_product_fast(self, product: str) -> List[Dict]:
        start_time = time.time()

        try:
            url = f"https://yandex.ru/finance/reviews/{product}"
            print(f"\n🌐 {product}")
            self.driver.get(url)
            time.sleep(3)  # Увеличенный таймаут для загрузки

            if not self.wait_for_reviews():
                print("  ❌ Отзывы не загрузились")
                return []

            scrolls = self.smart_scroll(product)
            reviews = self.extract_reviews_with_required_format(product)

            elapsed = time.time() - start_time
            print(f"  ⏱️ Обработано за {elapsed:.1f}с")

            return reviews

        except Exception as e:
            print(f"  ❌ Ошибка: {e}")
            return []

    def update_database(self, new_reviews: List[Dict]) -> bool:
        """Обновление базы данных в правильном формате"""
        if not new_reviews:
            return False

        try:
            if os.path.exists(self.database_file):
                df_existing = pd.read_parquet(self.database_file)
            else:
                df_existing = pd.DataFrame()

            df_new = pd.DataFrame(new_reviews)

            df_new['id'] = df_new['id'].astype('string')
            df_new['created_date'] = pd.to_datetime(df_new['created_date'])
            df_new['hub'] = df_new['hub'].astype('string')
            df_new['category_name'] = df_new['category_name'].astype('string')
            df_new['url'] = df_new['url'].astype('string')
            df_new['type'] = df_new['type'].astype('string')
            df_new['title'] = df_new['title'].astype('string')
            df_new['content'] = df_new['content'].astype('string')
            df_new['views_count'] = pd.to_numeric(df_new['views_count'], errors='coerce').astype('Int64')
            df_new['like_count'] = pd.to_numeric(df_new['like_count'], errors='coerce').astype('Int64')
            df_new['repost_count'] = pd.to_numeric(df_new['repost_count'], errors='coerce').astype('Int64')
            df_new['comments_count'] = pd.to_numeric(df_new['comments_count'], errors='coerce').astype('Int64')
            df_new['rating'] = pd.to_numeric(df_new['rating'], errors='coerce').astype('float64')

            if not df_existing.empty:
                df_merged = pd.concat([df_existing, df_new], ignore_index=True)
                df_merged = df_merged.drop_duplicates(subset=['id'], keep='first')
            else:
                df_merged = df_new

            columns_order = [
                'id', 'created_date', 'hub', 'category_name', 'url', 'type',
                'title', 'content', 'views_count', 'like_count', 'repost_count',
                'comments_count', 'rating'
            ]

            df_merged = df_merged[columns_order]
            df_merged.to_parquet(self.database_file, compression='snappy', index=False)

            print(f"\n💾 База обновлена: +{len(new_reviews)} записей, всего {len(df_merged)}")

            if self.use_cloud:
                print(f"☁️ Загрузка {self.database_file} в облако...")
                upload_to_cloud_sync(self.database_file)
            else:
                print(f"💾 Файл сохранен локально: {self.database_file}")

            return True

        except Exception as e:
            print(f"❌ Ошибка при обновлении базы: {e}")
            return False

    def batch_update(self, products: List[str]):
        """Пакетное обновление множества продуктов"""
        print(f"\n🚀 Быстрое обновление {len(products)} продуктов")
        print(f"   База данных: {self.database_file}")

        self.existing_ids, self.last_review_dates = self.load_database()

        all_new_reviews = []
        total_time = time.time()

        for i, product in enumerate(products, 1):
            print(f"\n[{i}/{len(products)}] ", end='')

            new_reviews = self.update_product_fast(product)

            if new_reviews:
                all_new_reviews.extend(new_reviews)

            if i < len(products):
                time.sleep(1)

        if all_new_reviews:
            self.update_database(all_new_reviews)

        elapsed = time.time() - total_time
        avg_time = elapsed / len(products)

        print(f"\n{'=' * 60}")
        print(f"✅ Обновление завершено!")
        print(f"   Общее время: {elapsed:.1f}с")
        print(f"   Среднее время на продукт: {avg_time:.1f}с")
        print(f"   Новых отзывов: {len(all_new_reviews)}")


def check_chrome_installed():
    """Проверка наличия Chrome и chromedriver"""
    import subprocess

    # Проверяем Chrome
    chrome_variants = ['google-chrome', 'chromium', 'chromium-browser', 'google-chrome-stable']
    chrome_found = False

    for variant in chrome_variants:
        try:
            subprocess.run([variant, '--version'],
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL,
                           check=True)
            chrome_found = True
            print(f"✅ Найден Chrome: {variant}")
            break
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue

    if not chrome_found:
        print("❌ Google Chrome не найден!")
        print("Установите Chrome:")
        print("  Ubuntu/Debian: sudo apt install google-chrome-stable")
        print("  или: sudo apt install chromium-browser")
        return False

    # Проверяем chromedriver
    try:
        subprocess.run(['chromedriver', '--version'],
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL,
                       check=True)
        print("✅ Найден chromedriver")
        return True
    except FileNotFoundError:
        print("❌ chromedriver не найден!")
        print("Установите chromedriver:")
        print("  sudo apt install chromium-chromedriver")
        print("  или скачайте с: https://chromedriver.chromium.org/")
        return False


if __name__ == "__main__":
    if not check_chrome_installed():
        print("\n❌ Установите необходимые компоненты и перезапустите скрипт")
        exit(1)
    parser = YandexFinanceSeleniumParser(headless=True, use_cloud=True)

    try:
        print("=" * 60)
        print("ПАРСЕР ОТЗЫВОВ ЯНДЕКС.ФИНАНСЫ")
        print("=" * 60)

        print("\nВыберите режим:")
        print("1. Быстрое обновление списка продуктов")
        print("2. Массовое обновление (50 продуктов)")
        print("3. Тест скорости на одном продукте")

        choice = input("\nВаш выбор (2-3): ").strip() or "2"

        if choice == "2":
            products = [
                "sberbank_deposit",
                "sberbank_credits",
                "sberbank_mortgage",
                "sberbank_debit-cards",
                "sberbank_credit-cards",
                "t-bank_deposit",
                "t-bank_credits",
                "t-bank_mortgage",
                "t-bank_debit-cards",
                "t-bank_credit-cards",
                "rosbank_deposit",
                "rosbank_credits",
                "rosbank_mortgage",
                "rosbank_debit-cards",
                "rosbank_credit-cards",
                "alfa-bank_deposit",
                "alfa-bank_credits",
                "alfa-bank_mortgage",
                "alfa-bank_debit-cards",
                "alfa-bank_credit-cards",
                "vtb_deposit",
                "vtb_credits",
                "vtb_mortgage",
                "vtb_debit-cards",
                "vtb_credit-cards",
                "mkb_deposit",
                "mkb_credits",
                "mkb_mortgage",
                "mkb_debit-cards",
                "mkb_credit-cards",
                "rosselkhozbank_deposit",
                "rosselkhozbank_credits",
                "rosselkhozbank_mortgage",
                "rosselkhozbank_debit-cards",
                "rosselkhozbank_credit-cards",
                "sovkombank_deposit",
                "sovkombank_credits",
                "sovkombank_mortgage",
                "sovkombank_debit-cards",
                "sovkombank_credit-cards",
                "bank-domrf_deposit",
                "bank-domrf_credits",
                "bank-domrf_mortgage",
                "bank-domrf_debit-cards",
                "bank-domrf_credit-cards",
                "gazprombank_deposit",
                "gazprombank_credits",
                "gazprombank_mortgage",
                "gazprombank_debit-cards",
                "gazprombank_credit-cards",
            ]

            print(f"\n⚡ Запуск массового обновления {len(products)} продуктов")
            confirm = input("Продолжить? (y/n): ").strip().lower()

            if confirm == 'y':
                parser.batch_update(products)

                print("\n" + "=" * 60)
                print("📊 ФИНАЛЬНАЯ ОБРАБОТКА")
                print("=" * 60)

                if os.path.exists(parser.database_file):
                    df = pd.read_parquet(parser.database_file)
                    excel_file = parser.database_file.replace('.parquet', '.xlsx')

                    print(f"📝 Создание Excel: {excel_file}")
                    excel_created = create_excel_with_retry(df, excel_file)

                    if excel_created:
                        if parser.use_cloud:
                            print(f"☁️ Загрузка Excel в облако...")
                            if upload_to_cloud_sync(excel_file):
                                delete_local_file(excel_file)
                                delete_local_file(parser.database_file)
                            else:
                                print(f"⚠️ Не удалось загрузить в облако, файлы сохранены локально")
                        else:
                            print(f"💾 Excel сохранен локально: {excel_file}")
                            print(f"💾 Parquet сохранен локально: {parser.database_file}")
                    else:
                        print(f"⚠️ Не удалось создать Excel, но Parquet сохранен: {parser.database_file}")

        elif choice == "3":
            product = input("ID продукта (Enter для sberbank_deposit): ").strip() or "sberbank_deposit"

            print(f"\n⚡ Тест скорости для {product}")

            start = time.time()
            parser.existing_ids, parser.last_review_dates = parser.load_database()
            reviews = parser.update_product_fast(product)
            elapsed = time.time() - start

            print(f"\n📊 Результаты теста:")
            print(f"   Время: {elapsed:.2f}с")
            print(f"   Новых отзывов: {len(reviews)}")

            if reviews:
                parser.update_database(reviews)

                if os.path.exists(parser.database_file):
                    df = pd.read_parquet(parser.database_file)
                    excel_file = parser.database_file.replace('.parquet', '.xlsx')

                    print(f"📝 Создание Excel: {excel_file}")
                    excel_created = create_excel_with_retry(df, excel_file)

                    if excel_created:
                        if parser.use_cloud:
                            print(f"☁️ Загрузка Excel в облако...")
                            if upload_to_cloud_sync(excel_file):
                                delete_local_file(excel_file)
                                delete_local_file(parser.database_file)
                            else:
                                print(f"⚠️ Не удалось загрузить в облако, файлы сохранены локально")
                        else:
                            print(f"💾 Excel сохранен локально: {excel_file}")
                            print(f"💾 Parquet сохранен локально: {parser.database_file}")
                    else:
                        print(f"⚠️ Не удалось создать Excel, но Parquet сохранен: {parser.database_file}")

    finally:
        parser.close()
        print("\n🔒 Браузер закрыт")
        if parser.use_cloud:
            print("✅ Все файлы загружены в облако и удалены локально")
        else:
            print("💾 Все файлы сохранены локально")