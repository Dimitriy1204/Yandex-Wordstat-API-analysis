"""Runtime hook: пути Prophet/CmdStan и __version__.py в собранном exe."""
import os
import sys

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    base = sys._MEIPASS
    prophet_dir = os.path.join(base, "prophet")
    os.makedirs(prophet_dir, exist_ok=True)
    ver_path = os.path.join(prophet_dir, "__version__.py")
    if not os.path.isfile(ver_path):
        with open(ver_path, "w", encoding="utf-8") as f:
            f.write('__version__ = "1.1.5"\n')
    stan = os.path.join(prophet_dir, "stan_model")
    if os.path.isdir(stan):
        os.environ["PROPHET_STAN_MODEL"] = stan
