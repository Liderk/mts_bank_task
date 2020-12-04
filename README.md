# mts_bank_task
## Парсинг данных с сайтов http://fssprus.ru/ и https://sudrf.ru/
Для работы скрипта fssp.py необходимо пройти [по ссылке](https://selenium-python.com/install-geckodriver)
и проследовать инструкции по установки geckodriver согласно вашей операционной системе.

Если необходимо(для Ubuntu можно не указывать) *executable_path* до установленного geckodriver,
расскоментив следующие строки в коде:
```
executable_path = указать путь до geckodriver в вашей системе
driver = Firefox(options=options, executable_path=executable_path
```

Установить *tesseract-ocr* согласно вашей ОС




Для ос ***windows 10*** можно воспользоваться инструкцией из этого [видео](https://www.youtube.com/watch?v=haHuVAUGY5Y).

При установке не забыть добавить русский язык



для  ***линукса***

sudo apt install tesseract-ocr

sudo apt install tesseract-ocr-rus

далее выполнить 
```
pip install -r requirements.txt
```

Для запуска скриптов выполнить

```
python3 sudrf.py
```
Результат работы можно посмотреть в файле *sudrf.xlsx*

```
python3 fssp.py
```

Результат работы можно посмотреть в файле *dossier.xlsx*