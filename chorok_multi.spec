# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None
app_dir = os.path.abspath('.')

# x64 VC++ 런타임 DLL 수동 번들.
# 빌드 머신이 Windows on ARM64일 때 System32에는 ARM64 DLL만 있으므로
# 반드시 x64 전용 경로(WindowsApps VCLibs.140.00.UWPDesktop)에서 가져와야 한다.
# 여기서 찾지 못하면 System32 폴백 (보통 순수 x64 머신)을 사용.
def _find_x64_vc_redist_dir():
    import glob
    patterns = [
        r'C:\Program Files\WindowsApps\Microsoft.VCLibs.140.00.UWPDesktop_*_x64__*',
    ]
    for pat in patterns:
        for p in sorted(glob.glob(pat), reverse=True):
            if os.path.exists(os.path.join(p, 'msvcp140.dll')):
                return p
    return None

vc_src = _find_x64_vc_redist_dir() or os.path.join(
    os.environ.get('SYSTEMROOT', r'C:\Windows'), 'System32'
)
vc_dll_names = (
    'msvcp140.dll', 'msvcp140_1.dll', 'msvcp140_2.dll',
    'vcruntime140.dll', 'vcruntime140_1.dll',
    'concrt140.dll',
)
msvcp_dlls = []
for name in vc_dll_names:
    p = os.path.join(vc_src, name)
    if os.path.exists(p):
        msvcp_dlls.append((p, '.'))

a = Analysis(
    ['main.py'],
    pathex=[app_dir],
    binaries=msvcp_dlls,
    datas=[
        (os.path.join(app_dir, 'data'), 'data'),
        (os.path.join(app_dir, 'sounds'), 'sounds'),
        (os.path.join(app_dir, 'green_auth'), 'green_auth'),
        (os.path.join(app_dir, 'bin'), 'bin'),
    ],
    hiddenimports=[
        'green_auth',
        'green_auth.auth_app',
        'green_auth.authenticator',
        'green_auth.config',
        'green_auth.credentials',
        'green_auth.login_dialog',
        'green_auth.screen_reader',
        'win32com.client',
        'lxml', 'lxml.etree',
        'bs4',
        'cryptography',
        'cryptography.fernet',
        'cryptography.hazmat',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='초록멀티 v1.3',
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon=os.path.join(app_dir, 'data', 'icon.ico'),
    version=os.path.join(app_dir, 'version_info.txt'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='초록멀티 v1.3',
)
