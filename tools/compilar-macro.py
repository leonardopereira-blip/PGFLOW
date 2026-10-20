"""Build PGFLOW VBA from reviewed source; no Office templates or p-code inputs.

Python >=3.10 and Git are required. Dependencies are pinned source checkouts,
kept outside the repository. Only source files listed below are checked out.

  python tools/compilar-macro.py --deps ../vba-source-tools --preparar-deps \
      --bin ../vba-source-tools/prototype/vbaProject-real.bin

Without --preparar-deps, existing clean checkouts at the pinned commits are
required. The default source is tools/RestaurarProtecaoOriginal.bas and the
default JS destination is MacroProtecaoVba.js. Repeated builds are byte-exact.
MS-OVBA writes project records and compressed source; pyOpenVBA writes the CFB
envelope and independently reads the finished project. Its bundled binary
templates are neither checked out nor loaded. Every module has MODULEOFFSET=0;
the only project cache is the documented seven-byte version-mismatch stub.
The generated VBA project is locked for viewing with the worksheet password;
reviewable VBA source remains in the repository. Password records follow
MS-OVBA sections 2.3.1.15-17 and 2.4.4 (native salted SHA-1 password hash).
"""

from pathlib import Path
from datetime import timedelta
import argparse
import base64
import hashlib
import json
import random
import re
import struct
import subprocess
import sys
import uuid


DEPS = {
    'MS-OVBA': ('Beakerboy/MS-OVBA', '2b8340dfb325df0560ae671fd78aeb19fcb6265d', ['/src/']),
    'MS-CFB': ('Beakerboy/MS-CFB', '870d4c18fde0e8164a8d5a4b26399640b761cae3', ['/src/ms_dtyp/filetime.py']),
    'MS-OVBA-Compression': ('Beakerboy/MS-OVBA-Compression', '41f50de6150e535ed7d523ad5bf507c5c717e2cd', ['/src/']),
    'MS-OVBA-Crypto': ('Beakerboy/MS-OVBA-Crypto', 'f49a8a8113f9e44744d10dbb7f530453dc3c0a6d', ['/src/']),
    'pyOpenVBA': ('WilliamSmithEdward/pyOpenVBA', '08c48d67c4a9d8b5905a6db0df8a61e5ab9f0fa4', [
        '/src/pyopenvba/cfb.py', '/src/pyopenvba/exceptions.py', '/src/pyopenvba/vba.py']),
}
PROJECT_PASSWORD = 'PGFLOW2026'
HASH_PROJECT_ID = '{00000000-0000-0000-0000-000000000000}'


def git(path, *args):
    return subprocess.check_output(['git', '-C', str(path), *args], text=True, stderr=subprocess.STDOUT).strip()


def dependencies(root, prepare):
    root.mkdir(parents=True, exist_ok=True)
    for name, (remote, commit, paths) in DEPS.items():
        folder = root / name
        if not folder.exists():
            if not prepare:
                raise RuntimeError(f'Missing source dependency {name}; use --preparar-deps')
            folder.mkdir()
            git(folder, 'init')
            git(folder, 'config', 'core.autocrlf', 'false')
            git(folder, 'remote', 'add', 'origin', f'https://github.com/{remote}.git')
            git(folder, 'sparse-checkout', 'init', '--no-cone')
            git(folder, 'sparse-checkout', 'set', '--no-cone', *paths)
            git(folder, 'fetch', '--depth=1', '--filter=blob:none', 'origin', commit)
            git(folder, 'checkout', '--detach', 'FETCH_HEAD')
        if git(folder, 'rev-parse', 'HEAD') != commit:
            raise RuntimeError(f'{name}: checkout differs from pinned commit {commit}')
        if git(folder, 'status', '--porcelain', '--untracked-files=no', '--', 'src'):
            raise RuntimeError(f'{name}: source checkout has local changes')
        sys.path.insert(0, str(folder / 'src'))


def document_source(name, base):
    return (f'Attribute VB_Name = "{name}"\r\nAttribute VB_Base = "0{{{base}}}"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')


def extract_source(data):
    """Independent small MS-OVBA reader, separate from both source encoders."""
    assert data[:1] == b'\x01'
    pos, output = 1, bytearray()
    while pos < len(data):
        start = pos
        header, = struct.unpack_from('<H', data, pos)
        assert (header >> 12) & 7 == 3
        end = start + (header & 0xFFF) + 3
        assert end <= len(data)
        pos += 2
        chunk_start = len(output)
        if not header & 0x8000:
            assert end - pos == 4096
            output.extend(data[pos:end])
            pos = end
            continue
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not flags & (1 << bit):
                    output.append(data[pos])
                    pos += 1
                else:
                    token, = struct.unpack_from('<H', data, pos)
                    pos += 2
                    width = max(4, (len(output) - chunk_start - 1).bit_length())
                    mask = 0xFFFF >> width
                    length = (token & mask) + 3
                    offset = (token >> (16 - width)) + 1
                    assert offset <= len(output) - chunk_start
                    for _ in range(length):
                        output.append(output[-offset])
        assert len(output) - chunk_start <= 4096
    return bytes(output)


def project_password_hash(password):
    """MS-OVBA 2.4.4.1/2/4; deterministic salt makes source builds repeatable."""
    key = b'PGFL'
    digest = hashlib.sha1(password.encode('cp1252') + key).digest()
    raw = key + digest
    mask = 0
    for value in raw:
        mask = (mask << 1) | bool(value)
    no_nulls = bytes(value or 1 for value in raw)
    return b'\xff' + mask.to_bytes(3, 'big') + no_nulls + b'\0'


def decode_project_record(hex_value, project_id):
    """Independent MS-OVBA 2.4.3.3 decoder for verification, not the writer."""
    data = bytes.fromhex(hex_value)
    seed, version_enc, key_enc = data[:3]
    assert version_enc ^ seed == 2
    plain_prior, encrypted_prior, encrypted_older = key_enc ^ seed, key_enc, version_enc
    assert plain_prior == sum(project_id.encode('cp1252')) & 255
    output = bytearray()
    for encrypted in data[3:]:
        plain = encrypted ^ ((plain_prior + encrypted_older) & 255)
        output.append(plain)
        plain_prior, encrypted_prior, encrypted_older = plain, encrypted, encrypted_prior
    ignored = (seed & 6) // 2
    length, = struct.unpack_from('<I', output, ignored)
    payload = bytes(output[ignored + 4:])
    assert len(payload) == length
    return payload


def validate_project_lock(stream, password):
    """Verify effective lock bits, hidden state and correct/incorrect passwords."""
    fields = dict(re.findall(r'^(ID|CMG|DPB|GC)="([^"]*)"', stream.decode('cp1252'), re.M))
    assert fields['ID'] == HASH_PROJECT_ID
    state = decode_project_record(fields['CMG'], fields['ID'])
    assert state == struct.pack('<I', 0x05)  # user + VBA editor, reserved bits zero
    assert decode_project_record(fields['GC'], fields['ID']) == b'\0'
    hashed = decode_project_record(fields['DPB'], fields['ID'])
    assert len(hashed) == 29 and hashed[0] == 0xFF and hashed[-1] == 0
    mask = int.from_bytes(hashed[1:4], 'big')
    raw = bytes(value if mask & (1 << (23 - i)) else 0 for i, value in enumerate(hashed[4:28]))
    key, expected = raw[:4], raw[4:]
    assert hashlib.sha1(password.encode('cp1252') + key).digest() == expected
    assert hashlib.sha1((password + '!').encode('cp1252') + key).digest() != expected


def build(source):
    from ms_ovba.vbaProject import VbaProject
    from ms_ovba.Models.Entities.doc_module import DocModule
    from ms_ovba.Models.Entities.std_module import StdModule
    from ms_ovba.Models.Entities.reference import Reference
    from ms_ovba.Models.Entities.reference_registered import ReferenceRegistered
    from ms_ovba.Models.Fields.libid_reference import LibidReference
    from ms_ovba.Views.dirStream import DirStream
    from ms_ovba.Views.project_view import ProjectView
    from ms_ovba.Views.project import Project
    from ms_ovba_compression.ms_ovba import MsOvba
    from ms_dtyp.filetime import Filetime
    from pyopenvba.cfb import CFB, DirEntry
    from pyopenvba.vba import compress, decompress, parse_vba_project

    # Upstream converts 1601 through a local Unix timestamp, invalid on Windows.
    # This date is unused here; constructing it directly keeps the adapter portable.
    Filetime.from_msfiletime = classmethod(lambda cls, ticks: cls(1601, 1, 1) + timedelta(microseconds=ticks // 10))
    class LockedProject(VbaProject):
        # Upstream convenience setters use reserved high bits and its hash
        # method is unfinished; emit the documented low-bit flags/hash here.
        @property
        def protection_state(self):
            return struct.pack('<I', 0x05)  # fUserProtected + fVBEProtected

        @property
        def visibility_state(self):
            return b'\0'

        @property
        def password(self):
            return project_password_hash(PROJECT_PASSWORD)

    project = LockedProject()
    project.project_id = HASH_PROJECT_ID  # required when DPB contains a hash
    project.add_attribute('VersionCompatible32', '393222000')
    for guid, version, path, name, short in (
        ('00020430-0000-0000-C000-000000000046', '2.0', r'C:\Windows\System32\stdole2.tlb', 'OLE Automation', 'stdole'),
        ('00020813-0000-0000-C000-000000000046', '1.9', r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE', 'Microsoft Excel 16.0 Object Library', 'Excel'),
    ):
        project.add_reference(Reference(ReferenceRegistered(LibidReference(uuid.UUID(guid), version, '0', path, name)), short))

    name_match = re.search(r'^Attribute VB_Name = "([A-Za-z_][A-Za-z0-9_]*)"', source, re.M)
    if not name_match:
        raise ValueError('Source must declare Attribute VB_Name')
    if re.search(r'(?im)^\s*(?:(?:Public|Private)\s+)?(?:Sub|Function)\s+(?:Auto_Open|Auto_Close|Workbook_\w+|Worksheet_\w+)\b', source):
        raise ValueError('Automatic execution/event procedures are outside this manual macro scope')
    modules = [
        (DocModule('ThisWorkbook'), document_source('ThisWorkbook', '00020819-0000-0000-C000-000000000046')),
        (StdModule(name_match.group(1)), source),
    ]
    # Source-authored root record and source streams only; no existing CFB is read.
    cfb = CFB(b'')
    root = DirEntry('Root Entry', 5, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFE, 0)
    root_raw = bytearray(cfb._synth_dir_entry_bytes(root))
    root_raw[67] = 1  # root is black per MS-CFB
    root.raw = bytes(root_raw)
    cfb._directory = [root]
    cfb.add_substorage_at((), 'VBA')
    codec = MsOvba()
    for module, text in modules:
        project.add_module(module)
        assert module.cache == b''
        # Upstream MS-OVBA-Compression mis-encodes multi-chunk sources. The
        # source-only pyOpenVBA encoder handles those; cross-check both readers.
        compressed = compress(text.encode('cp1252'))
        assert decompress(compressed).decode('cp1252') == text
        assert extract_source(compressed).decode('cp1252') == text
        cfb.add_stream_at(('VBA',), module.name, compressed)
    cfb.add_stream_at(('VBA',), 'dir', codec.compress(DirStream(project).to_bytes()))
    stub = ProjectView(project).to_bytes()
    assert stub == bytes.fromhex('CC61FFFF000300')
    cfb.add_stream_at(('VBA',), '_VBA_PROJECT', stub)
    random.seed(0)  # deterministic obfuscation of native project lock records
    cfb.add_stream_at((), 'PROJECT', Project(project).to_bytes())
    wm = b''.join(m.name.encode('cp1252') + b'\0' + m.name.encode('utf-16-le') + b'\0\0' for m, _ in modules) + b'\0\0'
    cfb.add_stream_at((), 'PROJECTwm', wm)
    binary = cfb.to_bytes()
    reread = CFB.from_bytes(binary)
    parsed = parse_vba_project(reread)
    assert parsed.module_names() == [m.name for m, _ in modules]
    assert not any(n.startswith('__SRP_') for n in reread.list_streams())
    assert reread.get_stream('_VBA_PROJECT') == stub
    validate_project_lock(reread.get_stream('PROJECT'), PROJECT_PASSWORD)
    for module, text in modules:
        extracted = parsed.get_module(module.name)
        assert extracted.source == text
        assert extracted.text_offset == 0
        assert extract_source(reread.get_stream(module.name)).decode('cp1252') == text
    return binary, [m.name for m, _ in modules]


def main():
    repo = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--deps', type=Path, required=True)
    parser.add_argument('--preparar-deps', action='store_true')
    parser.add_argument('--fonte', type=Path, default=repo / 'tools' / 'RestaurarProtecaoOriginal.bas')
    parser.add_argument('--bin', type=Path, required=True)
    parser.add_argument('--js', type=Path, default=repo / 'MacroProtecaoVba.js')
    args = parser.parse_args()
    dependencies(args.deps.resolve(), args.preparar_deps)
    raw_source = args.fonte.read_text(encoding='utf-8-sig')
    source = raw_source.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    source_hash = hashlib.sha256(source.replace('\r\n', '\n').encode('utf-8')).hexdigest()
    binary, modules = build(source)
    binary_hash = hashlib.sha256(binary).hexdigest()
    args.bin.parent.mkdir(parents=True, exist_ok=True)
    args.bin.write_bytes(binary)
    encoded = base64.b64encode(binary).decode('ascii')
    lines = [encoded[i:i + 120] for i in range(0, len(encoded), 120)]
    js = ('// VBA PGFLOW compilado exclusivamente da fonte tools/RestaurarProtecaoOriginal.bas.\n'
          '// Rebuild: tools/compilar-macro.py; dependencias de fonte com commits fixos.\n'
          f'// SHA-256 fonte (UTF-8, LF): {source_hash}\n'
          f'// SHA-256 vbaProject.bin: {binary_hash}\n'
          '// Sem eventos automaticos, templates externos, __SRP ou p-code.\n'
          '// Projeto VBA bloqueado para exibicao com a mesma senha das celulas.\n'
          f'var RE_MP_VBA_FONTE_SHA256 = "{source_hash}";\n'
          f'var RE_MP_VBA_BIN_SHA256 = "{binary_hash}";\n'
          'function _reSpMacroVba_() {\n  return [\n' +
          ',\n'.join('    "' + line + '"' for line in lines) + '\n  ].join(\'\');\n}\n')
    args.js.write_text(js, encoding='utf-8', newline='\n')
    print(json.dumps({'bytes': len(binary), 'sha256FonteLF': source_hash, 'sha256Bin': binary_hash,
                      'modules': modules, 'bin': str(args.bin), 'js': str(args.js)}, indent=2))


if __name__ == '__main__':
    main()
