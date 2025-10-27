import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import pandas as pd
import hashlib
from datetime import datetime, timedelta
import time
import re
import os
from typing import List, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import tqdm
from dotenv import load_dotenv
import boto3
from botocore.config import Config
from openpyxl import Workbook

# Загрузка переменных окружения
load_dotenv()

# Параметры облака
OBS_ACCESS_KEY = os.getenv("OBS_ACCESS_KEY")
OBS_SECRET_KEY = os.getenv("OBS_SECRET_KEY")
OBS_REGION = os.getenv("OBS_REGION")
OBS_ENDPOINT = os.getenv("OBS_ENDPOINT")
OBS_BUCKET = os.getenv("OBS_BUCKET")

required_vars = {
    "OBS_ACCESS_KEY": OBS_ACCESS_KEY,
    "OBS_SECRET_KEY": OBS_SECRET_KEY,
    "OBS_REGION": OBS_REGION,
    "OBS_ENDPOINT": OBS_ENDPOINT,
    "OBS_BUCKET": OBS_BUCKET,
}
for name, value in required_vars.items():
    if not value:
        raise EnvironmentError(f"Переменная окружения {name} не задана")


def upload_to_cloud_sync(file_path: str) -> bool:
    """Синхронная загрузка файла в облако"""
    try:
        s3_client = boto3.client(
            's3',
            aws_access_key_id=OBS_ACCESS_KEY,
            aws_secret_access_key=OBS_SECRET_KEY,
            region_name=OBS_REGION,
            endpoint_url=OBS_ENDPOINT,
            config=Config(signature_version='s3v4')
        )

        file_name = os.path.basename(file_path)

        s3_client.upload_file(
            file_path,
            OBS_BUCKET,
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


def download_latest_from_cloud(prefix: str) -> Optional[str]:
    """
    Скачивание последнего файла из облака по префиксу
    Возвращает имя скачанного файла или None
    """
    try:
        s3_client = boto3.client(
            's3',
            aws_access_key_id=OBS_ACCESS_KEY,
            aws_secret_access_key=OBS_SECRET_KEY,
            region_name=OBS_REGION,
            endpoint_url=OBS_ENDPOINT,
            config=Config(signature_version='s3v4')
        )

        response = s3_client.list_objects_v2(
            Bucket=OBS_BUCKET,
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

        local_filename = "temp_downloaded_findozor.parquet"
        s3_client.download_file(OBS_BUCKET, latest_file, local_filename)

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

            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"

            headers = list(df.columns)
            ws.append(headers)

            for _, row in df.iterrows():
                ws.append(row.tolist())

            widths = {
                'A': 20, 'B': 30, 'C': 12, 'D': 8, 'E': 8,
                'F': 20, 'G': 15, 'H': 18, 'I': 15, 'J': 10, 'K': 50, 'L': 50, 'M': 10
            }

            for col, width in widths.items():
                ws.column_dimensions[col].width = width

            wb.save(temp_filename)

            if os.path.exists(filename):
                os.remove(filename)
            os.rename(temp_filename, filename)

            print(f"✅ Excel создан: {filename}")
            return True

        except Exception as e:
            print(f"⚠️ Попытка {attempt + 1}/{max_retries} не удалась: {e}")
            time.sleep(1)

    return False

class FindozorTurboParser:
    def __init__(self, cutoff_date='2024-01-01', max_workers=8):
        self.base_url = "https://findozor.net"

        # Генерируем имя файла с датой
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.database_file = f"findozor_optimized_{current_date}.parquet"

        self.cutoff_date = datetime.strptime(cutoff_date, '%Y-%m-%d')
        self.max_workers = max_workers

        self.session = self._create_robust_session()

        print(f"📦 Целевой файл: {self.database_file}")

        # Скачивание последнего файла из облака
        print("🔍 Поиск последней версии базы данных в облаке...")
        downloaded_file = download_latest_from_cloud('findozor_optimized_')

        if downloaded_file:
            if os.path.exists(self.database_file):
                os.remove(self.database_file)
            os.rename(downloaded_file, self.database_file)
            print(f"✅ База данных готова к работе: {self.database_file}")
        else:
            print("📝 Будет создана новая база данных")

        self.existing_ids = self.load_existing_ids()

        print(f"📅 Фильтр: комментарии с {self.cutoff_date.strftime('%d.%m.%Y')} и новее")
        print(f"⚡️ Режим: {self.max_workers} потоков")
        print(f"🛡️ Устойчивость: включены повторные запросы при сбоях сети")

    def _create_robust_session(self) -> requests.Session:
        session = requests.Session()
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'ru-RU,ru;q=0.8'
        })
        retry_strategy = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    def load_existing_ids(self) -> set:
        if os.path.exists(self.database_file):
            try:
                df = pd.read_parquet(self.database_file)
                if 'id' in df.columns:
                    ids = set(df['id'].unique())
                    print(f"📂 Загружена база: {len(ids)} существующих записей")
                    return ids
            except Exception as e:
                print(f"⚠️ Ошибка загрузки базы: {e}")
        return set()

    def generate_id(self, content: str, url: str) -> str:
        return hashlib.md5((str(content) + str(url)).encode()).hexdigest()

    def parse_date(self, date_str: str) -> Optional[datetime]:
        if not date_str: return None
        try:
            if 'T' in date_str:
                return datetime.fromisoformat(date_str.replace('Z', '+00:00')).replace(tzinfo=None)
            date_text = date_str.strip().lower().replace(' в ', ' ')
            now = datetime.now()
            if 'сегодня' in date_text:
                time_match = re.search(r'(\d{1,2}):(\d{2})', date_text)
                return now.replace(hour=int(time_match.group(1)), minute=int(time_match.group(2)))
            if 'вчера' in date_text:
                yesterday = now - timedelta(days=1)
                time_match = re.search(r'(\d{1,2}):(\d{2})', date_text)
                return yesterday.replace(hour=int(time_match.group(1)), minute=int(time_match.group(2)))
            months = {'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'май': 5, 'июн': 6, 'июл': 7, 'авг': 8, 'сен': 9,
                      'окт': 10, 'ноя': 11, 'дек': 12}
            for month_name, month_num in months.items():
                if month_name in date_text:
                    match = re.search(r'(\d{1,2})\s+' + month_name + r'(?:\s+(\d{4}))?', date_text)
                    if match:
                        day, year = int(match.group(1)), int(match.group(2) or now.year)
                        return datetime(year, month_num, day)
        except:
            return None
        return datetime.now()

    def _parse_views(self, text: str) -> Optional[int]:
        text = text.lower().strip().replace(',', '')
        try:
            if 'k' in text: return int(float(text.replace('k', '')) * 1000)
            if 'm' in text: return int(float(text.replace('m', '')) * 1000000)
            return int(text)
        except:
            return None

    def get_all_sections(self) -> List[Dict]:
        print("\n📋 Шаг 1: Получение структуры форума...")
        try:
            response = self.session.get(f"{self.base_url}/forum/", timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            sections = []
            for block in soup.find_all('div', class_='block'):
                header = block.find('h2', class_='block-header')
                if not header or any(word in header.get_text().lower() for word in ['статистика', 'пользователи']):
                    continue
                for forum in block.find_all('h3', class_='node-title'):
                    link = forum.find('a')
                    if link and link.get('href'):
                        sections.append({'title': link.get_text(strip=True), 'url': self.base_url + link['href']})
            print(f"✅ Найдено {len(sections)} разделов.")
            return sections
        except Exception as e:
            print(f"❌ Ошибка получения разделов: {e}")
            return []

    def get_all_threads_from_section(self, section: Dict) -> List[Dict]:
        threads = []
        for page in range(1, 200):
            try:
                response = self.session.get(f"{section['url']}page-{page}", timeout=15)
                soup = BeautifulSoup(response.text, 'html.parser')
                items = soup.select('div.structItem')
                if not items: break
                for item in items:
                    title_elem = item.select_one('div.structItem-title a[href*="/threads/"]')
                    if not title_elem: continue
                    views_elem = item.select_one('dl.pairs--justified dt:-soup-contains("Просмотры") + dd')
                    views = self._parse_views(views_elem.get_text(strip=True)) if views_elem else None
                    threads.append({'title': title_elem.get_text(strip=True), 'url': self.base_url + title_elem['href'],
                                    'views_count': views})
                if not soup.select_one('a.pageNav-jump--next'): break
                time.sleep(0.1)
            except Exception:
                break
        return threads

    def get_thread_page_count(self, soup: BeautifulSoup) -> int:
        try:
            page_input = soup.find('input', class_='js-pageJumpPage')
            if page_input and 'max' in page_input.attrs and page_input['max'].isdigit():
                return int(page_input['max'])
            all_page_numbers = [int(match.group(1)) for link in soup.find_all('a', href=True) if
                                (match := re.search(r'/page-(\d+)', link['href']))]
            if all_page_numbers: return max(all_page_numbers)
            return 1
        except:
            return 1

    def _parse_thread(self, thread: Dict) -> List[Dict]:
        thread_posts = []
        stop_parsing_this_thread = False
        try:
            base_thread_url = re.sub(r'/page-\d+|/latest', '', thread['url']).strip('/')
            response = self.session.get(base_thread_url, timeout=30)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            total_pages = self.get_thread_page_count(soup)

            for page in range(total_pages, 0, -1):
                if page == 1:
                    page_soup = soup
                else:
                    page_url = f"{base_thread_url}/page-{page}"
                    response = self.session.get(page_url, timeout=30)
                    response.raise_for_status()
                    page_soup = BeautifulSoup(response.text, 'html.parser')

                posts_on_page = page_soup.find_all('article', class_='message')
                if not posts_on_page: continue

                for post in reversed(posts_on_page):
                    content_elem = post.find('div', class_='message-content')
                    if not content_elem: continue
                    for quote in content_elem.find_all('blockquote'): quote.decompose()
                    content = content_elem.get_text(strip=True)
                    if not content or len(content) < 10: continue

                    post_url_elem = post.find('a', class_='message-attribution-gadget')
                    post_url = self.base_url + (post_url_elem[
                                                    'href'] if post_url_elem and 'href' in post_url_elem.attrs else f"/threads/{thread['url'].split('/')[-2]}/")
                    record_id = self.generate_id(content, post_url)

                    if record_id in self.existing_ids:
                        tqdm.tqdm.write(
                            f"  [i] Остановка темы '{thread['title'][:40].ljust(43)}': найден существующий пост в базе на стр. {page}."
                        )
                        stop_parsing_this_thread = True
                        break

                    date_elem = post.find('time')
                    if not date_elem: continue
                    created_date = self.parse_date(date_elem.get('datetime') or date_elem.get_text())

                    if not created_date or created_date < self.cutoff_date:
                        tqdm.tqdm.write(
                            f"  [i] Остановка темы '{thread['title'][:40].ljust(43)}': найден старый пост ({created_date.strftime('%Y-%m-%d') if created_date else 'N/A'}) на стр. {page}."
                        )
                        stop_parsing_this_thread = True
                        break

                    thread_posts.append({
                        'id': record_id, 'created_date': created_date, 'hub': 'findozor.net/forum',
                        'category_name': 'Форум', 'url': post_url, 'type': 'форумное сообщение',
                        'title': thread['title'], 'content': content, 'views_count': thread.get('views_count'),
                        'like_count': sum(int(n) for n in re.findall(r'(\d+)', (
                                post.find('div', class_='reactionsBar') or BeautifulSoup('',
                                                                                         'html.parser')).get_text())) or None,
                        'repost_count': None, 'comments_count': None, 'rating': None
                    })

                if stop_parsing_this_thread:
                    break

                time.sleep(0.1)
        except Exception:
            pass
        return thread_posts

    def run_parser(self):
        start_time = time.time()
        sections = self.get_all_sections()
        if not sections: return

        print("\n📋 Шаг 2: Сбор всех тем форума (может занять несколько минут)...")
        all_threads = []
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = [executor.submit(self.get_all_threads_from_section, section) for section in sections]
            for future in tqdm.tqdm(as_completed(futures), total=len(sections), desc="Сбор тем"):
                all_threads.extend(future.result())

        print(f"✅ Всего найдено {len(all_threads)} тем для проверки.")
        print(f"\n💬 Шаг 3: Многопоточный парсинг комментариев (с {self.cutoff_date.strftime('%d.%m.%Y')})...")
        print(f"   ⚡️ Стратегия: парсинг С КОНЦА, с мгновенной остановкой по ID или дате.")

        results_from_threads = []
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = [executor.submit(self._parse_thread, thread) for thread in all_threads]
            for future in tqdm.tqdm(as_completed(futures), total=len(all_threads), desc="Парсинг тем"):
                results_from_threads.extend(future.result())

        print("\n💾 Шаг 4: Обработка и сохранение данных...")
        if results_from_threads:
            self.save_to_database(results_from_threads)
        else:
            print("ℹ️ Новых комментариев не найдено.")

        elapsed = time.time() - start_time
        print("\n" + "=" * 60)
        print("✅ ПАРСИНГ ЗАВЕРШЕН!")
        print(f"⏱️  Время выполнения: {timedelta(seconds=int(elapsed))}")
        if len(all_threads) > 0 and elapsed > 0:
            print(f"🚀  Средняя скорость: {len(all_threads) / elapsed:.1f} тем/сек")
        print("=" * 60)

    def save_to_database(self, posts: List[Dict]):
        try:
            df_new = pd.DataFrame(posts)
            print(f"   - Получено {len(df_new)} постов из потоков.")

            df_new.drop_duplicates(subset=['id'], keep='first', inplace=True)
            print(f"   - После удаления внутренних дублей: {len(df_new)} постов.")

            new_unique_posts = df_new[~df_new['id'].isin(self.existing_ids)]
            new_count = len(new_unique_posts)
            print(f"   - Реально новых для добавления: {new_count} постов.")

            if new_count == 0:
                print("   - Нет новых постов для сохранения.")
                return

            df_to_save = new_unique_posts
            if os.path.exists(self.database_file):
                df_existing = pd.read_parquet(self.database_file)
                df_to_save = pd.concat([df_existing, new_unique_posts], ignore_index=True)

            # Приводим типы данных
            for col in ['views_count', 'like_count', 'repost_count', 'comments_count']:
                df_to_save[col] = pd.to_numeric(df_to_save[col], errors='coerce').astype('Int64')
            df_to_save['rating'] = pd.to_numeric(df_to_save['rating'], errors='coerce').astype('float64')
            df_to_save['created_date'] = pd.to_datetime(df_to_save['created_date']).dt.tz_localize(None)

            # Порядок колонок
            columns_order = [
                'id', 'created_date', 'hub', 'category_name', 'url', 'type',
                'title', 'content', 'views_count', 'like_count', 'repost_count',
                'comments_count', 'rating'
            ]
            df_to_save = df_to_save[columns_order]

            # Сохраняем локально
            df_to_save.to_parquet(self.database_file, compression='snappy', index=False)
            print(f"✅ Успешно сохранено. Всего в базе: {len(df_to_save)} записей.")

            # Загружаем в облако
            print(f"☁️ Загрузка {self.database_file} в облако...")
            upload_to_cloud_sync(self.database_file)

        except Exception as e:
            print(f"❌ Ошибка сохранения: {e}")

if __name__ == "__main__":
    try:
        import pandas, pyarrow, tqdm
    except ImportError:
        print("Установка необходимых библиотек: pandas, pyarrow, tqdm...")
        os.system('pip install pandas pyarrow tqdm')
        print("Библиотеки установлены, пожалуйста, перезапустите скрипт.")
        exit()

    confirm = input("\nНачать полный парсинг? (y/n): ").strip().lower()
    if confirm == 'y':
        cutoff = input("Минимальная дата (YYYY-MM-DD, Enter = 2024-01-01): ").strip() or "2024-01-01"
        workers = input("Количество потоков (Enter = 8): ").strip() or "8"

        parser = FindozorTurboParser(cutoff_date=cutoff, max_workers=int(workers))
        parser.run_parser()

        # Финальная обработка - создание Excel
        print("\n" + "=" * 60)
        print("📊 ФИНАЛЬНАЯ ОБРАБОТКА")
        print("=" * 60)

        if os.path.exists(parser.database_file):
            df = pd.read_parquet(parser.database_file)
            excel_file = parser.database_file.replace('.parquet', '.xlsx')

            print(f"📝 Создание Excel: {excel_file}")
            if create_excel_with_retry(df, excel_file):
                print(f"☁️ Загрузка Excel в облако...")
                if upload_to_cloud_sync(excel_file):
                    delete_local_file(excel_file)

            # Удаляем parquet после всех операций
            delete_local_file(parser.database_file)

        print("\n✅ Все файлы загружены в облако и удалены локально")
    else:
        print("\n❌ Парсинг отменен.")