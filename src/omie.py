"""Cliente mínimo da API do Omie. Credenciais em config/credenciais.json (fora do git) ou OMIE_APP_KEY/OMIE_APP_SECRET."""
import json, os, subprocess, time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def _chave():
    if os.getenv("OMIE_APP_KEY"):
        return os.environ["OMIE_APP_KEY"], os.environ["OMIE_APP_SECRET"]
    c = json.loads((RAIZ / "config" / "credenciais.json").read_text())
    return c["app_key"], c["app_secret"]


def call(ep, metodo, param, tentativas=3):
    k, s = _chave()
    body = json.dumps({"call": metodo, "app_key": k, "app_secret": s, "param": [param]})
    for _ in range(tentativas):
        # ponytail: curl em vez de urllib — nesta máquina urllib falha por falta de CA bundle.
        r = subprocess.run(["curl", "-s", "--max-time", "120", "-X", "POST", f"https://app.omie.com.br/api/v1/{ep}/",
                            "-H", "Content-Type: application/json", "-d", body], capture_output=True, text=True)
        try:
            j = json.loads(r.stdout)
        except Exception:
            time.sleep(5); continue
        if "MISUSE_API_PROCESS" in r.stdout:
            raise SystemExit("Rate limit do Omie: aguarde ~30 min.")
        if "REDUNDANT" in r.stdout:
            time.sleep(26); continue
        return j
    raise RuntimeError(f"Omie não respondeu: {ep}/{metodo}")


def paginas(ep, metodo, chave_lista, base=None, por_pagina=100, pag="nPagina", rpp="nRegPorPagina", tot="nTotPaginas"):
    out, p = [], 1
    while True:
        prm = dict(base or {}); prm[pag] = p; prm[rpp] = por_pagina
        r = call(ep, metodo, prm)
        if chave_lista not in r:
            break
        out += r[chave_lista]
        if p >= int(r.get(tot, 1)):
            return out
        p += 1; time.sleep(1.2)
    return out
