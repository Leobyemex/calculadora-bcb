# -*- coding: utf-8 -*-
"""Monta o .zip da release a partir de dist\\CalculadoraBCB.

Uso:  python montar_zip_release.py 2.9.38

Conteúdo do zip:
  CalculadoraBCB/...                   -> app completo (exe + _internal)
  zz_ponte_atualizacao/CalculadoraBCB.exe
      -> PONTE para quem ainda está na 2.9.37 ou anterior. O atualizador
         antigo tinha um bug ('for /r' sem conferir o arquivo) e sempre copiava
         da ÚLTIMA pasta em ordem alfabética. Esta pasta, sem subpastas e com
         nome começando por 'zz', é justamente a última; assim o atualizador
         antigo troca o executável, e a versão nova completa a instalação
         (_internal) sozinha no primeiro uso. Quem instalar manualmente pode
         ignorar esta pasta.
"""
import os
import sys
import zipfile

def main():
    if len(sys.argv) < 2:
        print("uso: python montar_zip_release.py <versao>")
        return 1
    ver = sys.argv[1].strip().lstrip("v")
    src = os.path.join("dist", "CalculadoraBCB")
    exe = os.path.join(src, "CalculadoraBCB.exe")
    if not os.path.isfile(exe) or not os.path.isdir(os.path.join(src, "_internal")):
        print("ERRO: dist\\CalculadoraBCB (exe + _internal) nao encontrado.")
        return 1
    out = os.path.join("dist", f"CalculadoraBCB-v{ver}.zip")
    if os.path.exists(out):
        os.remove(out)
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for root, dirs, files in os.walk(src):
            dirs.sort()
            for f in sorted(files):
                full = os.path.join(root, f)
                arc = "CalculadoraBCB/" + os.path.relpath(full, src).replace(os.sep, "/")
                z.write(full, arc)
                n += 1
        # ponte: SEMPRE por último
        z.write(exe, "zz_ponte_atualizacao/CalculadoraBCB.exe")
    print(f"ZIP gerado: {out}  ({n} arquivos + ponte)")
    return 0

if __name__ == "__main__":
    sys.exit(main())
