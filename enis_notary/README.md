# Выгрузка нотариусов с enis.kz

Скачивает список нотариусов из раздела `https://enis.kz/Notary/NotaryByChamber/<ID>` в CSV и XLSX.

## Установка
```bash
pip install -r requirements.txt
```

## Запуск
```bash
python enis_notary.py                         # палата 4
python enis_notary.py --chambers 4 --details  # + данные из карточки каждого нотариуса
python enis_notary.py --chambers 1-20         # несколько палат в один файл
python enis_notary.py --chambers 4 --debug    # сохранить сырой HTML в ./debug_html
```

Параметры: `--delay` (пауза между запросами, по умолчанию 1 с), `--out` (имя файла).

Результат: `notaries_chamber_4.csv` (разделитель `;`, открывается в Excel) и `notaries_chamber_4.xlsx`.

Парсер сам определяет формат страницы (таблица или карточки), проходит пагинацию,
вытаскивает телефоны и email. Если выгрузка пустая или кривая — запустите с `--debug`
и пришлите файлы из `debug_html`, чтобы подогнать разбор под точную вёрстку.
