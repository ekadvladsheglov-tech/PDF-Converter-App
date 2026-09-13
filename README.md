# 📄 PDF Converter App

Простой, быстрый и полностью **локальный** десктопный конвертер PDF с современным интерфейсом.  
Никаких облаков, никакой отправки файлов на сервер — всё работает прямо на вашем ПК.

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

## ✨ Возможности

| Функция | Описание |
| :--- | :--- |
| 🔄 **PDF → Word** | Конвертация с сохранением разметки, таблиц и изображений |
| 🔄 **Word → PDF** | Быстрая конвертация .docx в PDF |
| ✂️ **Разделение PDF** | Разбивает многостраничный PDF на отдельные файлы |
| 🔗 **Объединение PDF** | Склеивает несколько PDF в один документ |

## 📥 Установка (для обычных пользователей)

1. Перейдите в раздел **[Releases](../../releases)**.
2. Скачайте последнюю версию **`PDF_Converter_Setup.exe`**.
3. Запустите установщик и следуйте инструкциям.

### ⚠️ Если ругается Windows SmartScreen

Программа бесплатная и без цифровой подписи, поэтому может появиться синее окно:

1. Нажмите **«Подробнее...»** (More info)
2. Затем **«Выполнить в любом случае»** (Run anyway)

Программа полностью безопасна и работает локально.

## 🛠 Сборка из исходников (для разработчиков)

### Автоматический способ (рекомендуется)

1. Установите [Python 3.10+](https://www.python.org/downloads/) (обязательно поставьте галочку **"Add Python to PATH"** при установке).
2. Установите [Inno Setup 6](https://jrsoftware.org/isinfo.php) в стандартную папку.
3. Скачайте репозиторий и запустите **`build_setup.bat`** двойным кликом.
4. Скрипт автоматически:
   - Создаст виртуальное окружение
   - Установит все зависимости (с правильной версией NumPy)
   - Соберет .exe через PyInstaller
   - Упакует всё в установщик через Inno Setup
5. Готовый `PDF_Converter_Setup.exe` появится на Рабочем столе.

### Ручной способ

```bash
git clone https://github.com/ekadvladsheglov-tech/PDF-Converter-App.git
cd PDF-Converter-App
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python pdf_converter.py