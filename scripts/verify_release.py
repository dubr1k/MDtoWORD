"""Validate exact version metadata and recursively inspect platform release ZIPs.

Uses only the standard library, so this also checks the downloaded public assets.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import plistlib
import re
import struct
import tomllib
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
PLATFORMS = ('macOS-arm64', 'Windows-x64')


def verify_metadata(tag: str) -> str:
    version = tomllib.loads((ROOT / 'pyproject.toml').read_text())['project']['version']
    assert tag == version, f'tag {tag} != version {version}'
    assert re.fullmatch(r'\d+\.\d+\.\d+', version), version
    spec = (ROOT / 'MDtoWORD.spec').read_text()
    assert f'"CFBundleShortVersionString": "{version}"' in spec
    windows = (ROOT / 'packaging/windows_version_info.txt').read_text()
    for key in ('FileVersion', 'ProductVersion'):
        assert f"StringStruct('{key}', '{version}')" in windows
    for key in ('filevers', 'prodvers'):
        expected = tuple(map(int, version.split('.'))) + (0,)
        actual = ast.literal_eval(re.search(rf'{key}=(\([^)]*\))', windows)[1])
        assert actual == expected, (key, actual, expected)
    notes = ROOT / 'docs' / 'releases' / f'RELEASE_NOTES_{tag}.md'
    assert notes.is_file() and len(notes.read_text().strip()) > 100
    print(f'Metadata and notes verified: {version}')
    return version


def verify_assets(directory: Path, version: str) -> None:
    expected = {f'MDtoWORD-{platform}.zip{suffix}' for platform in PLATFORMS for suffix in ('', '.sha256')}
    assert {p.name for p in directory.iterdir() if p.is_file()} == expected
    for platform in PLATFORMS:
        path = directory / f'MDtoWORD-{platform}.zip'
        checksum = path.with_suffix('.zip.sha256').read_text().strip().split()
        assert len(checksum) == 2 and checksum[1] == path.name
        with path.open('rb') as source:
            assert hashlib.file_digest(source, 'sha256').hexdigest() == checksum[0]
        with ZipFile(path) as archive:
            assert archive.testzip() is None, f'Corrupt ZIP: {path.name}'
            names = archive.namelist()
            assert len(names) == len(set(names)), 'Duplicate entries'
            for name in names:
                parts = PurePosixPath(name).parts
                assert '..' not in parts and not name.startswith('/'), name
                # Anchored first-party roots only: dependency src trees are legitimate.
                assert not any(p in ('.git', '.github', 'tests', '.env') for p in parts), name
                assert not any(p in ('AGENTS.md', 'pyproject.toml', 'MDtoWORD.spec') for p in parts), name
                assert not ('mdtoword' in parts and name.endswith(('.py', '.pyc'))), name
                assert not any(p in ('id_rsa', 'id_ed25519', 'credentials.json') for p in parts), name
            if platform == 'macOS-arm64':
                prefix = 'MDtoWORD.app/Contents/'
                executable = archive.read(prefix + 'MacOS/MDtoWORD')
                assert executable[:4] in (b'\xcf\xfa\xed\xfe', b'\xfe\xed\xfa\xcf')
                plist = plistlib.loads(archive.read(prefix + 'Info.plist'))
                assert plist['CFBundleShortVersionString'] == version
                icon = plist['CFBundleIconFile']
                if not icon.endswith('.icns'):
                    icon += '.icns'
                data = archive.read(prefix + 'Resources/' + icon)
                assert data[:4] == b'icns' and struct.unpack('>I', data[4:8])[0] == len(data)
                assert any(n.startswith(prefix + 'Frameworks/') and 'Python' in n for n in names)
                assert any(n.startswith(prefix + '_CodeSignature/') for n in names)
                assert any(n.endswith('/assets/macos-icon.png') for n in names)
            else:
                executable = archive.read('MDtoWORD/MDtoWORD.exe')
                assert executable[:2] == b'MZ'
                pe_offset = struct.unpack('<I', executable[60:64])[0]
                assert executable[pe_offset:pe_offset+4] == b'PE\0\0'
                assert version.encode('utf-16-le') in executable, 'Wrong Windows version resource'
                assert any(n.startswith('MDtoWORD/_internal/') and n.endswith('.dll') and '/python' in n.lower() for n in names)
                icons = [n for n in names if n.endswith('/assets/ico.png')]
                assert len(icons) == 1
                assert archive.read(icons[0]).startswith(b'\x89PNG\r\n\x1a\n')
            print(f'{path.name}: SHA-256, ZIP CRC, {len(names)} recursive entries, executable, runtime, icon and source/privacy checks verified')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--tag', required=True)
    parser.add_argument('--assets', type=Path)
    args = parser.parse_args()
    version = verify_metadata(args.tag)
    if args.assets:
        verify_assets(args.assets, version)
