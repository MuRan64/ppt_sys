# PyInstaller spec for Windows and macOS.
# Run from this release_build directory, or pass the absolute spec path.

from pathlib import Path
import os
import sys
from PyInstaller.utils.hooks import collect_submodules


# Resolve relative to this spec file so the build does not depend on the
# caller's current working directory.
ROOT = Path(SPECPATH).resolve().parent
TEST_MODULES = [
    "test_md_to_ppt",
    "test_sanbu_md_to_ppt",
    "test_english_md_to_ppt",
    "test_bibei_gushiwen_md_to_ppt",
    "test_wenyan_yuedu_jiafen_md_to_ppt",
    "test_zuhe_jiafen_md_to_ppt",
    "test_zhuji_100_md_to_ppt",
    "test_yilun_xinfangan_md_to_ppt",
]
BOOK_ADAPTERS = [
    "book_adapters.bibei_gushiwen",
    "book_adapters.wenyan_yuedu_jiafen",
    "book_adapters.zuhe_jiafen",
    "book_adapters.zhuji_100",
]

datas = [(str(ROOT / "static"), "static")]
hiddenimports = TEST_MODULES + BOOK_ADAPTERS + ["app_resources", "template_library", "generation_state"]
hiddenimports += collect_submodules("PIL")
TARGET_ARCH = os.environ.get("FBD_TARGET_ARCH") or None
EXTRA_BINARIES = []
if sys.platform == "win32":
    # Keep native dependencies from this Python environment together.  A
    # machine-wide PATH may contain incompatible copies from Poppler or other
    # applications, which PyInstaller can otherwise resolve first.
    conda_bin = Path(sys.prefix) / "Library" / "bin"
    for dll_name in (
        "libexpat.dll",
        "libcrypto-3-x64.dll",
        "libssl-3-x64.dll",
        "liblzma.dll",
        "LIBBZ2.dll",
    ):
        dll_path = conda_bin / dll_name
        if dll_path.is_file():
            EXTRA_BINARIES.append((str(dll_path), "."))

a = Analysis(
    [str(ROOT / "tools" / "fbd_md_gui.py")],
    pathex=[str(ROOT / "tools"), str(ROOT / "tests")],
    binaries=EXTRA_BINARIES,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy"],
    noarchive=False,
)
if sys.platform == "win32" and EXTRA_BINARIES:
    # Remove PATH-resolved copies before appending the environment-matched
    # DLLs, ensuring the collected package contains exactly one of each.
    pinned_names = {Path(item[0]).name.lower() for item in EXTRA_BINARIES}
    a.binaries = [item for item in a.binaries if item[0].lower() not in pinned_names]
    a.binaries.extend((Path(source).name, source, "BINARY") for source, _ in EXTRA_BINARIES)
pyz = PYZ(a.pure)
if sys.platform == "darwin":
    exe = EXE(
        pyz, a.scripts, a.binaries, a.datas, [], name="FbdToPpt",
        debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
        console=False,
        target_arch=TARGET_ARCH,
    )
    app = BUNDLE(exe, name="FbdToPpt.app", bundle_identifier="com.fbdtoppt.app")
else:
    exe = EXE(
        pyz, a.scripts, [], exclude_binaries=True, name="FbdToPpt",
        debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
        console=False,
    )
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="FbdToPpt")
