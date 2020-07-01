Конвертация Excel книги в JSON файл.

Зависимости
---
* pandas >= 1.0.5

Установка
---
```sh
$ virtualenv venv

# WINDOWS
$ venv\Scripts\activate
# LINUX
$ source venv/bin/activate

$ pip install -r requirements.txt 
```

Запуск
---
```sh
$ python xl_to_json.py <путь до excel файла> [-o <путь до выходного json файла>]
$ python xl_to_json.py data.xlsm
$ python xl_to_json.py data.xlsm -o result.json
```

Тестирование
---
```sh
$ cd test
$ pytest test.py
```
