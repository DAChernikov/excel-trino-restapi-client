# Install

В релизной ветке здесь должен лежать готовый Windows installer:

```text
install/setup.exe
```

Это основной пользовательский способ установки Trino Excel Client без Python, virtualenv и ручного запуска CLI.

Если файла `setup.exe` нет, его можно собрать на Windows:

```bat
pip install -e ".[test,package]"
python scripts\build_windows_exe.py --clean
python scripts\build_windows_installer.py
```

Сборка создаст `install\setup.exe`.
