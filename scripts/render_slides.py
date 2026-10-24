#!/usr/bin/env python3
"""Render every PPTX slide, verifying page coverage before publishing images."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

from pptx import Presentation


def powerpoint(deck: Path, dest: Path, width: int, count: int) -> None:
    import win32com.client

    app = win32com.client.DispatchEx('PowerPoint.Application')
    presentation = None
    try:
        presentation = app.Presentations.Open(str(deck), ReadOnly=True, Untitled=False, WithWindow=False)
        if presentation.Slides.Count != count:
            raise RuntimeError('PowerPoint slide count disagrees with PPTX structure')
        height = round(width * presentation.PageSetup.SlideHeight / presentation.PageSetup.SlideWidth)
        for n in range(1, count + 1):
            presentation.Slides(n).Export(str(dest / f'slide-{n:03d}.png'), 'PNG', width, height)
    finally:
        # This is an isolated application instance; never attach to or close a user session.
        try:
            if presentation is not None:
                presentation.Close()
        finally:
            app.Quit()


def libreoffice(deck: Path, dest: Path, width: int, count: int) -> None:
    import fitz

    exe = shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        for env in ('ProgramFiles', 'ProgramFiles(x86)'):
            base = os.environ.get(env)
            candidate = Path(base) / 'LibreOffice/program/soffice.exe' if base else None
            if candidate and candidate.is_file():
                exe = str(candidate)
                break
    if not exe:
        raise RuntimeError('LibreOffice executable not found')
    profile = dest / 'lo-profile'
    pdf_filter = 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}'
    result = subprocess.run(
        [exe, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--convert-to', pdf_filter,
         '--outdir', str(dest), str(deck)],
        capture_output=True, timeout=180, check=False,
    )
    pdf = dest / f'{deck.stem}.pdf'
    if result.returncode or not pdf.is_file():
        raise RuntimeError(f'LibreOffice PDF export failed: {result.stderr.decode(errors="replace")[:500]}')
    with fitz.open(pdf) as doc:
        if len(doc) != count:
            raise RuntimeError(f'PDF has {len(doc)} pages; PPTX has {count} slides')
        for n, page in enumerate(doc, 1):
            scale = width / page.rect.width
            page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).save(dest / f'slide-{n:03d}.png')


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('deck', type=Path)
    parser.add_argument('--output-dir', required=True, type=Path)
    parser.add_argument('--width', type=int, default=1600)
    args = parser.parse_args(argv)
    deck, output = args.deck.resolve(), args.output_dir.resolve()
    if not deck.is_file() or deck.suffix.lower() != '.pptx':
        parser.error('Provide an existing .pptx file')
    if args.width < 320:
        parser.error('--width must be at least 320')
    if output.exists() and any(output.iterdir()):
        parser.error('Use an empty output directory to avoid mixing stale renders')
    count = len(Presentation(deck).slides)
    if count == 0:
        parser.error('The presentation has no slides')
    backends = [('powerpoint', powerpoint)] if sys.platform == 'win32' else []
    backends.append(('libreoffice-pdf', libreoffice))
    errors = []
    for name, render in backends:
        # A fresh directory per attempt prevents partial results becoming false success.
        with tempfile.TemporaryDirectory(prefix='research-ppt-render-') as temp:
            staging = Path(temp)
            try:
                render(deck, staging, args.width, count)
                images = [staging / f'slide-{n:03d}.png' for n in range(1, count + 1)]
                if len(list(staging.glob('slide-*.png'))) != count:
                    raise RuntimeError('Rendered page count does not match PPTX')
                from PIL import Image
                for path in images:
                    with Image.open(path) as im:
                        im.verify()
            except Exception as exc:
                errors.append(f'{name}: {exc}')
                continue
            output.mkdir(parents=True, exist_ok=True)
            for path in images:
                shutil.copy2(path, output / path.name)
            manifest = dict(deck=str(deck), sha256=hashlib.sha256(deck.read_bytes()).hexdigest(),
                            renderer=name, expected_slides=count, rendered_slides=len(images),
                            images=[path.name for path in images], fallback_errors=errors)
            (output / 'render_manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
            print(f'Rendered {count}/{count} slides with {name}: {output}')
            return 0
    sys.stderr.write('Full-deck rendering failed.\n' + '\n'.join(errors) + '\n')
    return 2


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
