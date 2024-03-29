# D2PC

**D**OCX **to** **P**DF **c**onverter

## Описание

D2PC (DOCX to PDF converter) - программа для конвертации *.docx* файлов в формат *.pdf* через контекстное меню Windows
(Если говорить проще, то принцип такой: **ПКМ по DOCX файлу → Конвертировать в PDF → Получить PDF**).

Подробнее об установке, использовании, переустановке, удалении и самостоятельной сборке написано ниже.

**⭐ Не забудь поставить звезду репозиторию ⭐**

![](readme_image_1.png)

## Как установить?

### Скачивание файла

Перейди в раздел **Releases** в репозитории и найди последний релиз.
К релизу будет прикреплён файл **D2PC.exe**, его нужно скачать.

### Проблемы с антивирусом

При скачивании файла браузер и антивирус будут ругаться. Файл будет помечен как опасный, так как он прописывает себя
в реестр, чтобы его можно было вызывать из контекстного меню. Необходимо добавить файл в исключения антивируса.

*(Если вы, как и я, параноик, то внизу описано, как exe-файл можно собрать вручную из исходного кода)*

После скачивания перемести файл в папку, в которой он будет храниться
(о перемещении файла уже после установки написано ниже, в пункте *"Как переустановить?"*).

Кликни по перемещённому файлу правой кнопкой мыши и запусти от имени администратора. После этого файл
пропишется в реестр и D2PC можно будет использовать.

## Как использовать?

| Windows 10                                   | Windows 11                                                                    |
|----------------------------------------------|-------------------------------------------------------------------------------|
| 1. Кликни правой кнопкой мыши по .docx файлу | 1. Кликни правой кнопкой мыши по .docx файлу                                  |
| 2. В открывшемся меню снизу найди пункт D2PC | 2. В открывшемся меню снизу найди пункт **Показать дополнительные параметры** |
| 3. Нажми на пункт **D2PC**                   | 3. Нажми на пункт **Показать дополнительные параметры*                        |
|                                              | 4. В открывшемся меню снизу найди пункт **D2PC**                              |
|                                              | 4. Нажми на пункт **D2PC**                                                    |

*(Подсказка: в Windows 11 можно сразу открыть полное меню, кликнув по файлу правой кнопкой мыши с зажатым Shift)*

## Как переустановить?

Если необходимо поменять местоположение файла **D2PC.exe**, то всё просто: просто перемести его в нужное место,
и запусти от имени администратора, как это было при установке.

## Как удалить?

D2PC не предусматривает автоматического удаления. Если его всё же необходимо удалить, то нужно проделать следующие шаги:

1. Открыть редактор реестра
2. В **HKEY_CLASSES_ROOT** необходимо найти раздел с названием **.docx**
3. Нужно открыть найденный раздел и посмотреть его значение по умолчанию, обычно это **"Word.Document.12"**
4. В **HKEY_CLASSES_ROOT** необходимо найти раздел с названием, полученным на предыдущем шаге
   *(опять же, обычно это раздел **"Word.Document.12"**)*
5. Раскройте найденный раздел. Внутри него должен быть раздел **shell**, его тоже необходимо раскрыть
6. Внутри раздела **shell** будет раздел **D2PC**, по нему нужно кликнуть правой кнопкой мыши и выбрать удаление
   *(Внимание, раздел **shell** удалять НЕЛЬЗЯ!)*
7. Осталось только удалить файл **D2PC.exe**, на этом удаление закончено

## Как собрать?

Чтобы собрать exe-файл самостоятельно, нужно:

1. Создать виртуальную среду python (Virtualenv) в папке .venv в корне проекта
2. Установить в виртуальную среду следующие библиотеки: **docx2pdf**, **pyinstaller** и **Pillow**.
   Это можно сделать командой `.\.venv\Scripts\pip.exe install docx2pdf pyinstaller Pillow`
   *(команду нужно запускать из корня проекта)*
3. Запустить файл build.bat
