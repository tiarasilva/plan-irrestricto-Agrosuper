# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('holidays')


block_cipher = None


a = Analysis(
    ['generar_plan_irrestricto.py'],
    pathex=[],
    binaries=[],
    datas=[('Colaboraciones plan de ventas/Asignación venta.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Fechas de zarpe - Logística.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Maestro de materiales.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Parametros.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Pedidos Planta-Puerto-Embarcado.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Pedidos Stock.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Producción.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Proyeccion Plan de Venta.xlsx', 'Colaboraciones plan de ventas'), ('Colaboraciones plan de ventas/Volumen por contenedor.xlsx', 'Colaboraciones plan de ventas'), ('Img/Notice.png', 'Img')],
    hiddenimports=hiddenimports,
    hookspath=['.'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='generar_plan_irrestricto',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Img/ico.icns:Img'],
)
