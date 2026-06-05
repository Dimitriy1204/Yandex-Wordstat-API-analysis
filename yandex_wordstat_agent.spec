# -*- mode: python ; coding: utf-8 -*-
"""
Spec-файл для сборки Yandex Wordstat Analysis Agent в один .exe.
"""
import os
import sys
import importlib.util

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

PROJECT_DIR = os.getcwd()
sys.path.insert(0, PROJECT_DIR)

# ── Prophet / CmdStan ─────────────────────────────────────────────
_prophet_bins = []
_prophet_py_datas = []
_prophet_spec = importlib.util.find_spec("prophet")
if _prophet_spec and _prophet_spec.origin:
    _PROPHET_DIR = os.path.dirname(_prophet_spec.origin)
    for _root, _dirs, _files in os.walk(_PROPHET_DIR):
        for _fn in _files:
            if _fn.endswith((".py", ".pyi", ".typed", ".json", ".bin", ".exe", ".dll", ".so")):
                _src = os.path.join(_root, _fn)
                _rel = os.path.relpath(_src, _PROPHET_DIR)
                _dest = os.path.join("prophet", os.path.dirname(_rel)).replace("\\", "/")
                _prophet_py_datas.append((_src, _dest or "prophet"))
    _STAN_MODEL_DIR = os.path.join(_PROPHET_DIR, "stan_model")
    _CMDSTAN_BIN_DIR = os.path.join(_STAN_MODEL_DIR, "cmdstan-2.37.0", "bin")
    _TBB_DIR = os.path.join(
        _STAN_MODEL_DIR, "cmdstan-2.37.0", "stan", "lib", "stan_math", "lib", "tbb",
    )
    _bundles = [
        (os.path.join(_STAN_MODEL_DIR, "prophet_model.bin"), os.path.join("prophet", "stan_model")),
        (os.path.join(_CMDSTAN_BIN_DIR, "stanc.exe"), os.path.join("prophet", "stan_model", "cmdstan-2.37.0", "bin")),
        (os.path.join(_CMDSTAN_BIN_DIR, "diagnose.exe"), os.path.join("prophet", "stan_model", "cmdstan-2.37.0", "bin")),
        (os.path.join(_CMDSTAN_BIN_DIR, "print.exe"), os.path.join("prophet", "stan_model", "cmdstan-2.37.0", "bin")),
        (os.path.join(_CMDSTAN_BIN_DIR, "stansummary.exe"), os.path.join("prophet", "stan_model", "cmdstan-2.37.0", "bin")),
        (os.path.join(_TBB_DIR, "tbb.dll"), os.path.join("prophet", "stan_model", "cmdstan-2.37.0", "stan", "lib", "stan_math", "lib", "tbb")),
    ]
    for src, dest in _bundles:
        if os.path.isfile(src):
            _prophet_bins.append((src, dest))

prophet_datas = collect_data_files("prophet")
prophet_hidden = collect_submodules("prophet")

# ── Datas: UI + образец Excel (если есть) ─────────────────────────
_datas = [
    (os.path.join(PROJECT_DIR, "app", "templates", "index.html"), "app/templates"),
]
for _tpl_name in ("yandex_analysis_exemple.xlsx", "yandex_analysis_example.xlsx"):
    _tpl_path = os.path.join(PROJECT_DIR, _tpl_name)
    if os.path.isfile(_tpl_path):
        _datas.append((_tpl_path, "."))
        print(f"Excel template embedded in exe: {_tpl_name}")
        break
else:
    print("WARNING: yandex_analysis_exemple.xlsx not found — build exe without Excel template.")
_datas += prophet_datas + _prophet_py_datas

block_cipher = None

a = Analysis(
    ["run.py"],
    pathex=[PROJECT_DIR],
    binaries=_prophet_bins,
    datas=_datas,
    hiddenimports=[
        "fastapi", "uvicorn", "uvicorn.logging", "uvicorn.loops", "uvicorn.loops.auto",
        "uvicorn.protocols", "uvicorn.protocols.http", "uvicorn.protocols.http.auto",
        "uvicorn.protocols.websockets", "uvicorn.protocols.websockets.auto",
        "uvicorn.middleware", "uvicorn.middleware.proxy_headers", "uvicorn.server", "uvicorn.config",
        "starlette", "starlette.templating", "starlette.staticfiles",
        "jinja2", "jinja2.ext",
        "requests", "urllib3",
        "pandas", "pandas._libs", "numpy", "numpy.core._multiarray_umath",
        "xlsxwriter", "openpyxl",
        "app.chart_data", "app.excel_emoji_headers",
        "app.yandexgpt_prompts", "app.yandexgpt_client",
        "cryptography", "cryptography.fernet",
        "cryptography.hazmat.primitives", "cryptography.hazmat.primitives.kdf",
        "cryptography.hazmat.primitives.kdf.pbkdf2",
        "cryptography.hazmat.backends", "cryptography.hazmat.backends.openssl",
        "socks", "sockshandler",
        "pydantic", "pydantic.dataclasses",
        "prophet", "prophet.forecaster", "prophet.models", "prophet.__version__",
        "cmdstanpy", "cmdstanpy.cmdstan_path", "cmdstanpy.utils",
        "holidays", "tqdm", "stanio", "convertdate", "lunarcalendar",
        "scipy", "scipy.stats", "scipy.special", "scipy.linalg",
        "app.excel_header_template", "app.prophet_forecast",
        "python_multipart", "dotenv",
        "http.cookies", "http.client",
        "email.mime.multipart", "email.mime.base", "email.mime.text", "email.encoders",
        "asyncio", "concurrent", "concurrent.futures",
    ] + prophet_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[os.path.join(PROJECT_DIR, "pyi_rth_prophet.py")],
    excludes=[
        "tkinter", "test", "pdb", "doctest", "nose", "pytest",
        "sphinx", "setuptools", "pip", "wheel",
        "IPython", "jupyter", "jupyter_client", "notebook", "nbformat", "nbconvert",
        "bokeh", "plotly", "sympy", "curses",
        "PyQt5", "PyQt6", "PySide2", "PySide6", "cv2", "opencv", "pydantic.v1",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="yandex_wordstat_agent",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
